import asyncio
import configparser
import logging
import os
import smtplib
import sys
from time import sleep
import socket
import win32serviceutil
import servicemanager
import win32event
import win32service
import traceback
import selectors

from aiosmtpd.controller import Controller

__version__ = '1.0.3'

class SMTPProxy(win32serviceutil.ServiceFramework):
    _svc_name_ = 'SMTPProxy'
    _svc_display_name_ = 'SMTPProxy'
    _svc_description_ = "SMTPProxy adapt non-TLS SMTP email to TLS SMTP email"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        log_path = os.path.join(sys.path[0], 'smtpproxy.log')
        logging.basicConfig(filename=log_path, level=logging.INFO)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

        socket.setdefaulttimeout(60)
        self.isAlive = True

    def SvcStop(self):
        logging.debug("stop")
        self.isAlive = False
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        logging.debug("run")
        self.isAlive = True
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTING, (self._svc_name_, ''))
        try:
            self.main()
        except Exception as e:
            logging.error(traceback.print_stack())
            logging.error(str(e))

        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def main(self):
        logging.debug("SMPTProxy starting")

        config_path = os.path.join(
            sys.path[0],
            'config.ini'
        )
        logging.info(f"Loading Config: {config_path}")
        if not os.path.exists(config_path):
            raise Exception("Config file not found: {}".format(config_path))

        config = configparser.ConfigParser()
        config.read(config_path)

        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ": Listening on %s:%s" % (
                                  config.get('local', 'host', fallback='127.0.0.1'),
                                  config.get('local', 'port', fallback=25))))

        use_auth = config.getboolean('remote', 'smtp_auth', fallback=False)
        if use_auth:
            auth = {
                'user': config.get('remote', 'smtp_auth_user'),
                'password': config.get('remote', 'smtp_auth_password')
            }
        else:
            auth = None

        selector = selectors.SelectSelector()
        loop = asyncio.SelectorEventLoop(selector)

        logging.debug("Loading controller")
        controller = Controller(
            MailProxyHandler(
                host=config.get('remote', 'host'),
                port=config.getint('remote', 'port', fallback=25),
                auth=auth,
                use_ssl=config.getboolean('remote', 'use_ssl', fallback=False),
                starttls=config.getboolean('remote', 'starttls', fallback=False),
            ),
            loop=loop,
            hostname=config.get('local', 'host', fallback='127.0.0.1'),
            port=config.getint('local', 'port', fallback=25)
        )
        controller.loop.set_debug(True)
        controller.start()
        logging.debug("started controller")
        while controller.loop.is_running():
            if not self.isAlive:
                controller.stop()
            sleep(0.5)

        logging.debug("Finished")


class MailProxyHandler:
    def __init__(self, host, port=0, auth=None, use_ssl=False, starttls=False):
        self._host = host
        self._port = port
        auth = auth or {}
        self._auth_user = auth.get('user')
        self._auth_password = auth.get('password')
        self._use_ssl = use_ssl
        self._starttls = starttls

    async def handle_DATA(self, server, session, envelope):
        try:
            refused = self._deliver(envelope)
        except smtplib.SMTPRecipientsRefused as e:
            logging.info('Got SMTPRecipientsRefused: %s', refused)
            return "553 Recipients refused {}".format(' '.join(refused.keys()))
        except smtplib.SMTPResponseException as e:
            return "{} {}".format(e.smtp_code, e.smtp_error)
        else:
            if refused:
                logging.info('Recipients refused: %s', refused)
            return '250 OK'

    # adapted from https://github.com/aio-libs/aiosmtpd/blob/master/aiosmtpd/handlers.py
    def _deliver(self, envelope):
        refused = {}
        try:
            if self._use_ssl:
                s = smtplib.SMTP_SSL(self._host, self._port)
            else:
                s = smtplib.SMTP(self._host, self._port)
            s.connect(self._host, self._port)
            if self._starttls:
                s.starttls()
                s.ehlo()
            if self._auth_user and self._auth_password:
                s.login(self._auth_user, self._auth_password)
            try:
                refused = s.sendmail(
                    envelope.mail_from,
                    envelope.rcpt_tos,
                    envelope.original_content
                )
            finally:
                s.quit()
        except (OSError, smtplib.SMTPException) as e:
            logging.exception('got %s', e.__class__)
            # All recipients were refused. If the exception had an associated
            # error code, use it.  Otherwise, fake it with a SMTP 554 status code. 
            errcode = getattr(e, 'smtp_code', 554)
            errmsg = getattr(e, 'smtp_error', e.__class__)
            raise smtplib.SMTPResponseException(errcode, errmsg.decode())


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SMTPProxy)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SMTPProxy)

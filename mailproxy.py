import asyncio
import configparser
import logging
import os
import smtplib
import sys
from time import sleep

from aiosmtpd.controller import Controller

# From  http://thepythoncorner.com/dev/how-to-create-a-windows-service-in-python/

__version__ = '1.0.3'

'''
From http://thepythoncorner.com/dev/how-to-create-a-windows-service-in-python/

SMWinservice
by Davide Mastromatteo

Base class to create winservice in Python
-----------------------------------------

Instructions:

1. Just create a new class that inherits from this base class
2. Define into the new class the variables
   _svc_name_ = "nameOfWinservice"
   _svc_display_name_ = "name of the Winservice that will be displayed in scm"
   _svc_description_ = "description of the Winservice that will be displayed in scm"
3. Override the three main methods:
    def start(self) : if you need to do something at the service initialization.
                      A good idea is to put here the inizialization of the running condition
    def stop(self)  : if you need to do something just before the service is stopped.
                      A good idea is to put here the invalidation of the running condition
    def main(self)  : your actual run loop. Just create a loop based on your running condition
4. Define the entry point of your module calling the method "parse_command_line" of the new class
5. Enjoy
'''

import socket

import win32serviceutil

import servicemanager
import win32event
import win32service
import logging


class SMTPProxy(win32serviceutil.ServiceFramework):
    _svc_name_ = 'SMTPProxy'
    _svc_display_name_ = 'SMTPProxy'
    _svc_description_ = "SMTPProxy adapt non-TLS SMTP email to TLS SMTP email"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        logging.debug("__init__")
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
                              servicemanager.PYS_SERVICE_STARTED, (self._svc_name_, ''))
        self.main()
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def main(self):
        logging.debug(":main")
        logging.debug("SMPTProxy starting")
        while True:
            logging.debug("Sleeping")
            sleep(2)
        if len(sys.argv) == 2:
            config_path = sys.argv[1]
        else:
            config_path = os.path.join(
                sys.path[0],
                'config.ini'
            )
        logging.debug(f"Config: ${config_path}")
        if not os.path.exists(config_path):
            raise Exception("Config file not found: {}".format(config_path))

        config = configparser.ConfigParser()
        config.read(config_path)

        use_auth = config.getboolean('remote', 'smtp_auth', fallback=False)
        if use_auth:
            auth = {
                'user': config.get('remote', 'smtp_auth_user'),
                'password': config.get('remote', 'smtp_auth_password')
            }
        else:
            auth = None

        controller = Controller(
            MailProxyHandler(
                host=config.get('remote', 'host'),
                port=config.getint('remote', 'port', fallback=25),
                auth=auth,
                use_ssl=config.getboolean('remote', 'use_ssl', fallback=False),
                starttls=config.getboolean('remote', 'starttls', fallback=False),
            ),
            hostname=config.get('local', 'host', fallback='127.0.0.1'),
            port=config.getint('local', 'port', fallback=25)
        )
        controller.loop.set_debug(True)
        controller.start()
        while controller.loop.is_running():
            if not self.isrunning:
                controller.stop()
            sleep(0.2)


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
    logging.basicConfig(filename='D:\\temp\\smtpproxy.log', level=logging.DEBUG)
    logging.debug("Starting")
    logging.debug(f"Num args = ${sys.argv}")
    # pywin32_postinstall.py -install

    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SMTPProxy)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SMTPProxy)

    logging.debug("Done")

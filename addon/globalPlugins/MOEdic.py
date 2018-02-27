# -*- coding: UTF-8 -*-
from __future__ import unicode_literals
from logHandler import log
from subprocess import PIPE, Popen
from threading import Thread
import errno
import globalPluginHandler
import os
import re
import sys
import ui

dir = os.path.dirname(__file__)
exe_base = os.path.join(dir, "MOEdic")

def record(fd, output_fn):
    for line in iter(fd.readline, ""):
        try:
            line = line.decode("UTF-8")
            output_fn(line.rstrip())
        except:
            log.exception(sys.exc_info()[0])

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

    def __init__(self):
        super(GlobalPlugin, self).__init__()
        self.subproc = None
        self.workers = []

    def terminate(self):
        try:
            if self.subproc is not None:
                self.subproc.terminate()
                for w in self.workers:
                    w.join()
        except:
            pass
        super(GlobalPlugin, self).terminate()

    def script_MOEdic(self, gesture):
        if self.subproc is not None and self.subproc.poll() is None:
            ui.message(_("已經在查詢了！"))
            return
        log.info(_("開始新的一次查詢。"))
        for w in self.workers:
            w.join()
        self.workers = [None] * 2
        try:
            self.subproc = Popen([os.extsep.join([exe_base + "-x64", "exe"]), "/quick"], stdout=PIPE, stderr=PIPE)
        except WindowsError as e:
            if e.errno != errno.ENOEXEC:
                raise e
            self.subproc = Popen([os.extsep.join([exe_base, "exe"]), "/quick"], stdout=PIPE, stderr=PIPE)
        self.workers[0] = Thread(name="MOEdic-stdout", target=record, args=(self.subproc.stdout, log.info))
        self.workers[1] = Thread(name="MOEdic-stderr", target=record, args=(self.subproc.stderr, ui.message))
        for w in self.workers:
            w.daemon = True
            w.start()

    script_MOEdic.__doc__ = _("國字快查")

    __gestures = { "kb:windows+numpad2": "MOEdic" }

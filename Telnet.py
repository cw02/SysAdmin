#!/usr/bin/env python
# -*- coding: ASCII -*-

"""Foobar.py: Description of what foobar does."""

__version__ = "0.0.2"

__author__      = "Christopher Weller"
__copyright__   ="Copyright 2009, Planet Earth"
__credits__ = ["Christopher Weller"]
__license__ = "GPL"
__contact__="Christopher Weller"
__deprecated__="False"
__date__="13 April 2017"
__maintainer__ = "Christopher Weller"
__email__ = "christopher.weller@gm.com"
__status__ = "Production"

import getpass, sys, telnetlib

HOST = "<HOST_IP>"
user = raw_input("Enter your remote account: ")
password = getpass.getpass()

tn = telnetlib.Telnet(HOST)

tn.read_until("login: ")
tn.write(user + "\r\n")
if password:
   tn.read_until("Password: ")
   tn.write(password + "\r\n")

tn.write("vt100\r\n") 
tn.write("ls\r\n")
tn.write("exit\r\n")
print (tn.read_all())
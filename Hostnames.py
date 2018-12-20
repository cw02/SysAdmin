#!/usr/bin/python
from netaddr import IPNetwork
import socket
import netifaces as nif
 
 
 
for loop_ip in ['10.36.96.0', '10.36.96.255']:
 
    try:
        #Will get the hostname and IP address.
            dns = socket.gethostbyaddr(loop_ip)
            hostnm = dns[0]
            ipadd  = (", ".join(dns[2]))
            print hostnm.ljust(10), ipadd.rjust(20)
         
 
    except socket.error, msg:
            print msg
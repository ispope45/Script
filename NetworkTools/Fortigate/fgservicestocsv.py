#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# This file is part of fgpoliciestocsv.
#
# Copyright (C) 2014, 2020, Thomas Debize <tdebize at mail.com>
# All rights reserved.
#
# fgpoliciestocsv is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# fgpoliciestocsv is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with fgpoliciestocsv.  If not, see <http://www.gnu.org/licenses/>.

from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

from codecs import open
from os import path 
import sys
import re
import csv
import os

# OptionParser imports
from optparse import OptionParser
from optparse import OptionGroup

# Options definition
parser = OptionParser(usage="%prog [options]")

main_grp = OptionGroup(parser, 'Main parameters')
main_grp.add_option('-i', '--input-file', help='Partial or full Fortigate configuration file. Ex: fgfw.cfg')
main_grp.add_option('-o', '--output-file', help='Output csv file (default ./services-out.csv)', default=path.abspath(path.join(os.getcwd(), './services-out.csv')))
main_grp.add_option('-s', '--skip-header', help='Do not print the csv header', action='store_true', default=False)
main_grp.add_option('-n', '--newline', help='Insert a newline between each group for better readability', action='store_true', default=False)
main_grp.add_option('-d', '--delimiter', help='CSV delimiter (default ";")', default=';')
parser.option_groups.extend([main_grp])

# Python 2 and 3 compatibility
if (sys.version_info < (3, 0)):
    fd_read_options = 'rb'
    fd_write_options = 'wb'
else:
    fd_read_options = 'r'
    fd_write_options = 'w'

# Handful patterns
# -- Entering group definition block
p_entering_service_block = re.compile(r'^\s*config firewall service ', re.IGNORECASE)

# -- Exiting group definition block
p_exiting_service_block = re.compile(r'^end$', re.IGNORECASE)

# -- Commiting the current group definition and going to the next one
p_service_next = re.compile(r'^next$', re.IGNORECASE)

# -- Policy number
p_service_name = re.compile(r'^\s*edit\s+"(?P<service_name>.*)"$', re.IGNORECASE)

# -- Policy setting
p_service_set = re.compile(r'^\s*set\s+(?P<service_key>\S+)\s+(?P<service_value>.*)$', re.IGNORECASE)

# Functions
def parse(options):
    """
        Parse the data according to several regexes
        
        @param options:  options
        @rtype: return a list of groups ( [ {'id' : '1', 'srcintf' : 'internal', ...}, {'id' : '2', 'srcintf' : 'external', ...}, ... ] )  
                and the list of unique seen keys ['id', 'srcintf', 'dstintf', ...]
    """
    global p_entering_service_block, p_exiting_service_block, p_service_next, p_service_name, p_service_set
    
    in_service_block = False
    
    service_list = []
    service_elem = {}
    
    order_keys = []
    
    with open(options.input_file, mode=fd_read_options) as fd_input:
        for line in fd_input:
            line = line.strip()
            
            # We match a group block
            if p_entering_service_block.search(line):
                in_service_block = True
            
            # We are in a group block
            if in_service_block:
                if p_service_name.search(line):
                    service_name = p_service_name.search(line).group('service_name')
                    service_elem['name'] = service_name
                    if not('name' in order_keys):
                        order_keys.append('name')
                
                # We match a setting
                if p_service_set.search(line):
                    service_key = p_service_set.search(line).group('service_key')
                    if not(service_key in order_keys):
                        order_keys.append(service_key)
                    
                    service_value = p_service_set.search(line).group('service_value').strip()
                    service_value = re.sub('["]', '', service_value)
                    
                    service_elem[service_key] = service_value
                
                # We are done with the current group id
                if p_service_next.search(line):
                    service_list.append(service_elem)
                    service_elem = {}
            
            # We are exiting the group block
            if p_exiting_service_block.search(line):
                in_service_block = False
    
    return (service_list, order_keys)


def generate_csv(results, keys, options):
    """
        Generate a plain ';' separated csv file
    """
    if results and keys:
        with open(options.output_file, mode=fd_write_options) as fd_output:
            spamwriter = csv.writer(fd_output, delimiter=options.delimiter, quoting=csv.QUOTE_ALL, lineterminator='\n')
            
            if not(options.skip_header):
                spamwriter.writerow(keys)
            
            for group in results:
                output_line = []
                
                for key in keys:
                    if key in group.keys():
                        if "member" == key:
                            output_line.append("|".join(group[key].split(" ")))
                        else:
                            output_line.append(group[key])
                    else:
                        output_line.append('')
            
                spamwriter.writerow(output_line)
                if options.newline:
                    spamwriter.writerow('')
        
        fd_output.close()
    
    return

def main():
    """
        Dat main
    """
    global parser
    
    options, arguments = parser.parse_args()
    
    if (options.input_file == None):
        parser.error('Please specify a valid input file')
                
    results, keys = parse(options)
    generate_csv(results, keys, options)
    
    return None

if __name__ == "__main__" :
    main()
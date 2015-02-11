#!/usr/bin/env ruby

require 'optparse'
require 'csv'

NEWLINE = "\r"

def error(usage, message=nil)
  puts message if message
  puts usage
  exit(-1) 
end

# configure command line options
options = {}
optparse = OptionParser.new do|opts|
  opts.banner = "cms_exchanger: translate CMS export files to Office 365 format
  Usage: top_exchanger inputfile [outputfile]"
  opts.on( '-h', '--help', 'Display this screen' ) do
     puts opts
     exit
  end
end
optparse.parse!
# Check if files were specified
if ARGV.empty?
  error(optparse, nil)
end
export_file = ARGV.shift # this is the the export file we're going to translate
unless File.file?(export_file)
  error(optparse, "#{export_file} doesn't exist")
end
import_file = ARGV.shift || "O365_import_file.csv" # this is the translated file we're going to ouput

# get headers from the import template
import_headers = CSV.open("O365_import_template.csv").first
# open output file & add headers
output = CSV.open(import_file, "wb")
output << import_headers

# this csv has mappings between cms -> exhange fields
mapping = CSV.open('field_mapping.csv', headers: true).read

# try to guess text encoding. it's usually iso-8859-1 or us-ascii
encoding = `file -b --mime-encoding #{export_file}`.chomp
export = File.open(export_file, encoding: "#{encoding}:UTF-8")
export_headers = export.readline(NEWLINE).chomp.split(',') # have to specify newline char since it is not \n
export.each_line(NEWLINE) do |line|
  # not only is there a \r to get rid of, but there is an extra comma for some reason!
  # regex escapes "" that are nested within quoted fields like so: ""nested""
  escaped = line.chomp.chop.gsub(/(?<!^|,)"(?!,|$)/,'""')
  next if escaped.empty? # these crappy files have empty lines for some reason
  contact = CSV::Row.new( export_headers, CSV.parse_line(escaped) )
  new_contact = CSV::Row.new(import_headers, [])
  mapping.each do |column|
    new_contact[ column["Exchange Column"] ] = contact[ column['CMS DB Column'] ].to_s.strip # strip fields with nothing but blank space
  end
  output << new_contact
end





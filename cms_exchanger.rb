#!/usr/bin/env ruby

require 'optparse'
require 'csv'
require 'active_support/core_ext/hash/slice'

# configure command line options
# more will go here eventually
options = {}
optparse = OptionParser.new do|opts|
  opts.banner = "cms_exchanger: translate CMS export files to Office 365 format\n
  Usage: top_exchanger inputfile [outputfile]"
  opts.on( '-h', '--help', 'Display this screen' ) do
     puts opts
     exit
  end
end
optparse.parse!
# Check if files were specified
if ARGV.empty?
  puts optparse
  exit(-1)
end
export_file = ARGV.shift # this is the the export file we're going to translate
import_file = ARGV.shift # this is the translated file we're going to ouput
import_file = "O365_import_file.csv" if !import_file

# get headers from the import template
import_headers = CSV.open("O365_import_template.csv").first
# open output file & add headers
output = CSV.open(import_file, "wb")
output << import_headers

# this csv has mappings between cms -> exhange fields
mapping = CSV.open('field_mapping.csv', headers: true).read

CSV.foreach(export_file, encoding: "ISO-8859-1:UTF-8", headers: true) do |contact|
  new_contact = CSV::Row.new(import_headers, [])
  mapping.each do |column|
    new_contact[ column["Exchange Column"] ] = "#{ contact[ column['CMS DB Column'] ] }".strip # get rid of fields with nothing but blank space
  end
  output << new_contact
end

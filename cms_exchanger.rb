#!/usr/bin/env ruby

require 'optparse'
require 'csv'
require 'active_support/core_ext/hash/slice'
require 'yaml'

NEWLINE = "\r"
mapping = YAML.load_file("mapping.yml")
# get headers for exchange import file
IMPORT_HEADERS = mapping.keys

def error(usage, message)
  puts message if message
  puts usage
  exit(-1) 
end

def scrub_csv(line)
  # not only is there a \r to get rid of, but there is an extra comma for some reason! hend chop.comp()
  # regex escapes "" that are nested within quoted fields so they are double instead like: ""nested""
  line.chomp.chop.gsub(/(?<!^|,)"(?!,|$)/,'""')
end

def street_address(fields)
  fields.delete_if {|f| f == "0"} # hsn is sometimes 0
  fields.join(' ').strip.squeeze(' ')
end

def zip_code(fields)
  fields[1].empty? ? fields[0] : fields.join("-") # they will always be strings, no Nil values
end

def concat_fields(method, fields)
  # use key string from hash as name of method to call
  send(method, fields)
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
# check if files were specified, handle errors
error(optparse, nil) if ARGV.empty?
export_file = ARGV.shift # this is the file we're going to translate
error(optparse, "#{export_file} doesn't exist") unless File.file?(export_file)

# this is the translated file we're going to ouput, give it a name if not specified
import_file = ARGV.shift || "O365_import_file.csv"
output = CSV.open(import_file, "w")
output << IMPORT_HEADERS

# try to guess text encoding. it's usually iso-8859-1 or us-ascii
encoding = `file -b --mime-encoding #{export_file}`.chomp
# this has to be opened as a plain text file instead of csv because it may
# have unescaped quotes that need to be dealt with before it can be parsed
export = File.open(export_file, encoding: "#{encoding}:UTF-8")
EXPORT_HEADERS = export.readline(NEWLINE).chomp.split(',') # have to specify newline char since it is not \n

# map the fields
export.each_line(NEWLINE) do |line|
  scrubbed = scrub_csv(line)
  next if scrubbed.empty? # these csv files have empty lines at the end sometimes
  contact = CSV::Row.new( EXPORT_HEADERS, CSV.parse_line(scrubbed) )
  new_contact = CSV::Row.new( IMPORT_HEADERS, [] )
  mapping.each do |exchange_col, cms_col|
    if cms_col.kind_of?(Hash)
      cms_col.each do |method, fields|
        new_contact[exchange_col] = concat_fields( method, contact.to_hash.slice(*fields).values )
      end
    else
      new_contact[exchange_col] = contact[cms_col].to_s.strip # strip the fields with nothing but blank spaces
    end
  end
  output << new_contact
end

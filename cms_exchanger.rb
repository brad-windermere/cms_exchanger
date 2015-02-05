#!/usr/bin/env ruby

require 'optparse'
require 'csv'
require 'active_support/core_ext/hash/slice'

# o365 doesn't care about header case, so "Business street", "Home street", & "Web page" are fine
# "E-mail Address" would need to be changed to "E-mail" though 
unique_to_o365 = ["Business street", "Home street", "E-mail", "Web page"]

# unique_to_cms = ["Business Street",
#  "Business Street 3",
#  "Home Street",
#  "Home Street 3",
#  "Other Street 3",
#  "Car Phone",
#  "ISDN",
#  "Account",
#  "Anniversary",
#  "Billing Information",
#  "Birthday",
#  "Directory Server",
#  "E-mail Address",
#  "Gender",
#  "Government ID Number",
#  "Hobby",
#  "Initials",
#  "Internet Free Busy",
#  "Keywords",
#  "Language",
#  "Location",
#  "Mileage",
#  "Organizational ID Number",
#  "Priority",
#  "Private",
#  "Profession",
#  "Referred By",
#  "Sensitivity",
#  "Spouse",
#  "User 1",
#  "User 2",
#  "User 3",
#  "User 4",
#  "Web Page"]

# test this out without changing "Business street", "Home street", or "Web page"
# will need to change "E-mail Address" to "Email" once done running
# test to see if the case matters for those columns (it shouldn't)
unique_to_cms = ["Business Street 3",
 "Home Street 3",
 "Other Street 3",
 "Car Phone",
 "ISDN",
 "Account",
 "Anniversary",
 "Billing Information",
 "Birthday",
 "Directory Server",
 "Gender",
 "Government ID Number",
 "Hobby",
 "Initials",
 "Internet Free Busy",
 "Keywords",
 "Language",
 "Location",
 "Mileage",
 "Organizational ID Number",
 "Priority",
 "Private",
 "Profession",
 "Referred By",
 "Sensitivity",
 "Spouse",
 "User 1",
 "User 2",
 "User 3",
 "User 4"]

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

CSV.foreach(export_file, encoding: "ISO-8859-1:UTF-8", headers: true) do |contact|
  # new_contact = CSV::Row.new(import_headers, [])
  contact.delete_if {|header, field| unique_to_cms.include? header}
  output << contact
end

#!/usr/bin/env ruby

require 'optparse'
require 'csv'
require 'active_support/core_ext/hash/slice'

# o365 doesn't care about header case, so "Business street", "Home street", & "Web page" are fine
# "E-mail Address" would need to be changed to "E-mail" though 
unique_to_o365 = ["Business street", "Home street", "E-mail", "Web page"]

unique_to_cms = ["Business Street",
 "Business Street 3",
 "Home Street",
 "Home Street 3",
 "Other Street 3",
 "Car Phone",
 "ISDN",
 "Account",
 "Anniversary",
 "Billing Information",
 "Birthday",
 "Directory Server",
 "E-mail Address",
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
 "User 4",
 "Web Page"]

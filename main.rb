require 'xsv'

def get_org(org_string, org_zip)
  return org_zip if org_zip

  if org_string.instance_of?(String) && org_string.include?('For shipments originating in')
    # puts org_string
    org_string.scan(/[0-9]{3}/)
  end
end

def get_zones(row, zones_matches)
  if row[0].instance_of?(String) && row[0].include?('the following Postal Codes are Zone')
    matches = row[0].scan(/Zone ([0-9]+)/)
  end
  matches || zones_matches
end

def parse_file(file, output_file)
  x = Xsv.open(file) # => #<Xsv::Workbook sheets=1>

  sheet = x.sheets[0]
  org_zip = nil
  zones_matches = nil

  sheet.each do |row|
    next if row[0].nil? || row[0] == ''

    org_zip = get_org(row[0], org_zip)

    if row[0] && row[0].instance_of?(String) && row[0]&.strip&.match?(/^[0-9]{3}$/)
      output_file.puts("#{org_zip[0]}, #{org_zip[1]}, #{row[0]}, #{row[1]}, #{row[2]}, #{row[3]}, #{row[4]}, #{row[5]}, #{row[6]}")
    end

    zones_matches = get_zones(row, zones_matches)

    next unless row[0] && row[0].to_s.match?(/^[0-9]{5}$/)

    next if zones_matches.empty?

    row.each do |c|
      if c
        output_file.puts("#{c}, #{c}, #{zones_matches[0][0]}, '', #{zones_matches[1][0]}, '', '', '', #{zones_matches[2][0]}")
      end
    end
  end
end

def init
  filenames = Dir.entries('zone_files')
  output_file = File.open('output.csv', 'w')
  output_file.puts('Send Range Start,Send Range End,Receive Range Start,Receive Range End,Ground,3 Day Select,2nd Day Air,2nd Day Air AM,Next Day Air Saver,Next Day Air')
  filenames.each do |f|
    puts f
    parse_file("zone_files/#{f}", output_file) if f.include?('.xlsx')
  end
  output_file.close
end

init

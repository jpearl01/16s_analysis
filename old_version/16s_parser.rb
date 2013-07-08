#!/usr/bin/env ruby

require 'axlsx'
require 'trollop'

#The idea behind this program is to be able to pass hierarchy files from the MultiClassifier.jar program 
#which identifies where 16s sequences are from and create excel files with graphs of the data.
#Usage: 
#Setup options hash
opts = Trollop::options do
  opt :hier_file, "MultiClassfier.jar output (class_hier.txt)", type: :string, short: '-f'
  opt :neg_control, "Optional negative control file (also class_hier.txt)", type: :string, short: '-n'
end

class Organism
  attr_accessor :taxid, :lineage, :name, :rank, :number
end

all_organisms = []
file = File.basename(opts[:hier_file])
samp_file = File.open(opts[:hier_file], 'r')

classifications = %w(phylum class order family genus species)
num_classes = {}
classifications.each do |c|
  num_classes["#{c}"] = 0
end

samp_file.lines do |line|
  next unless samp_file.lineno > 3
  org = Organism.new
  fields = line.split("\s")
  org.taxid = fields[0]
  org.lineage = fields[1]
  org.name = fields[2].gsub(/"/,"")
  org.rank = fields[3]
  org.number = fields[4]
  num_classes[org.rank] = 1 + num_classes[org.rank] if  num_classes.has_key?(org.rank)
  all_organisms.push(org)
end

#Sort the organisms by how many there are
all_organisms.sort! { |a,b| a.number.to_i <=> b.number.to_i }

#Add the first workbook, this holds the original annotation data for each cluster
p = Axlsx::Package.new
wb = p.workbook

#Build the sheets needed for this project into the workbook
classifications.each_entry do |i|
  wb.add_worksheet(:name => "#{i}") do |sheet|
    sheet.add_row ["All bacteria classified as #{i}"]
    sheet.add_row %w(name number)
  end
end

#The sort worked, but I want it from largest to smallest, add all the data to the worksheet
all_organisms.reverse_each do |o|
  next unless classifications.include?(o.rank)
  wb.sheet_by_name("#{o.rank}").add_row [o.name, o.number]
end

classifications.each_entry do |i|
  data_end = num_classes[i].to_s
  next if data_end.eql?("0")
  sheet = wb.sheet_by_name("#{i}")
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "D6", :end_at => "U40") do |chart|
    chart.add_series  :labels => sheet["A3:A#{data_end}"], :data => sheet["B3:B#{data_end}"], :title => sheet["A1"]
    #chart.catAxis.label_rotation = -45
  end
end

p.serialize("#{file}.xlsx")

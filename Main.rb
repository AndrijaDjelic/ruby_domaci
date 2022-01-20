require 'roo'
require './xlsx'

run = Xlsx.new('test.xlsx')

puts "Ispis po headerima"
p run.findHeaders
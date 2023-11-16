require'./meta.rb'

session = GoogleDrive::Session.from_config("config.json")
worksheet = session.spreadsheet_by_key("1CPmguO2YewOl5Ggs1DCpclPGHrPsbOTn6IVBcJwK60M").worksheets[0]
sheet = GoogleSheetTable.new(worksheet)

# Primeri upotrebe
sheet['prva kolona'][2] = 1
puts "Vrednosti reda: #{sheet.row(3)}"
puts "Kolona: #{sheet['Prva kolona']}"
sheet.each { |cell| puts cell }
p sheet['prva kolona'][2]

puts "Suma: #{sheet.prvakolona.sum}"
puts "Prosek: #{sheet.prvakolona.avg}"
puts sheet.prvakolona
puts "Red: #{sheet.indeks('rn2310')}"

puts sheet.prvakolona.map { |cell| cell.to_i + 1}
puts sheet.prvakolona.select { |cell| cell.to_i.odd?}
puts sheet.prvakolona.reduce(0) { |sum, cell| sum + cell.to_i }

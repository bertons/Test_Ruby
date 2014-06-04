require 'win32ole'
connection = WIN32OLE.new('ADODB.Connection')
connection.Open('Provider=Microsoft.Jet.OLEDB.4.0;
                 Data Source=database_test_ruby.mdb')

i = 0
file = File.new("Dictionnaire.csv", "r")
while (line = file.gets)
  line = line.split(/"/)
  word = line[1]
  definition = line[3]
  if definition.length > 255
    definition = definition[0...255]
  end
  connection.Execute("INSERT INTO Dictionnaire (Words, Definitions) VALUES (\"#{word.to_s}\", \"#{definition.to_s}\");")
  i = i + 1
end
file.close
connection.Close

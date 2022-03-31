#!/usr/bin/python3
from imp import source_from_cache
import sys, getopt
import csv
from terrainclassification import TerrainClassification

def parseFile (iFile, format = "pretty"):
   ofile = "temp.csv"
   s = None
   e = None

   f = open(iFile, 'r')
   labelList = f.readlines()


   e= open("errors.txt", 'w')

   outCSV = open(ofile, 'w')

   if format == "csv":
      labelWriter = csv.writer(outCSV, delimiter=',')
      labelWriter.writerow(labelList[0].split(','))
   elif format == "pretty":
      s = open(ofile, 'w')
   else:
      print (format + " is not supported. Valid options are 'csv' or 'pretty'")
   #Write header rows


   count = 0
   errors = []

   for i in labelList[1:]:
      parse = None
      count +=1

      try:
         parse = TerrainClassification(i.split(",")[3])
         # parse = TerrainClassification(i[1])
         print("\nLine Number: " +str(count) + parse.printPretty())
         if format == "csv":
            print (parse.printCSV())
            labelWriter.writerow(parse.printCSV())
         elif format == "pretty":
            s.write ("\n\nLine Number: " +str(count) + parse.printPretty())

      except Exception as a:
         print (a)
         print ("Error occured at line: " +str(count))
         print ("Problem string: " + parse.terrainString)

         errors.append(parse.terrainString)
         e.write(parse.terrainString + "\n")

   print ("Total records: " + str(len(labelList)))
   print ("Processing Errors: " + str(len(errors)))

   f.close()
   if s:
      s.close()
   e.close()
   if outCSV:
      outCSV.close()

def parseString(string, format = "pretty"):

   x = TerrainClassification(string)
   if format == "pretty":
      print(x.printPretty())
   elif format == "csv":
      print (x.printCSV())
   else:
      print("Error unspecified format. \nValid options are: pretty or csv.")

def main(argv):
   scriptname = "parseLabel.py"
   inputfile = ''
   successfile = ''
   string = ''
   format= "pretty"

   try:
      opts, args = getopt.getopt(argv,"h:i:o:s:f:",["ifile=","ofile=","string=", "format="])
   except getopt.GetoptError:
      print (scriptname+ ' -i <inputfile> -o <outputfile>')
      sys.exit(2)


   for opt, arg in opts:
      if opt == '-h':
         print (scriptname +
                  '\n -i \t<inputfile> \n'
                  ' -o \t<outputfile> \n'
                  ' -s \t<string>')
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputfile = arg
      elif opt in ("-o", "--ofile"):
         outputfile = arg
      elif opt in ("-s", "--string"):
         string = arg
      elif opt in ("-f", "--format"):
         format = arg

   # if not outputfile:
   #    print("all")
   if inputfile:
      parseFile(inputfile, format)
   elif string:
      parseString(string, format)



   # print ('Output file is "', successfile)

# if __name__ == "__main__":
main(sys.argv[1:])


# string = r'sgFGptM.dsmMbM\xsCv\cLGpM-XsV'

# string2 = r'sgFGptM.dsmMbM/xsCv\\zcLGpM-XsV'

# string3 = r'sgFGptM.dsmMbM//xsCv\zcLGpM-RdsbXsVI'
# print (string3)

# # x = TerrainClassification(string)
# # # y = TerrainClassification(string2)
# z = TerrainClassification(string3)
# print(z.printPretty())


# print (z.printCSV())

# print (sys.argv)
#!/usr/bin/python3
import sys, getopt
from terrainclassification import TerrainClassification

def parseFile (filename, format = "pretty"):

   f = open(filename, 'r')
   labelList = f.readlines()

   e= open("errors.txt", 'w')
   s = open("output.csv", 'w')

   count = 0
   errors = []
   problem = 0

   if problem:

   # print (labelList[problem])
      parse = TerrainClassification(labelList[problem].split(",")[3])
      print (parse.printPretty())

   else:
      for i in labelList[1:]:
         parse = None
         count +=1

         try:
            parse = TerrainClassification(i.split(",")[3])
            # parse = TerrainClassification(i[1])
            print("\nLine Number: " +str(count) + parse.printPretty())
            s.write ("\n\nLine Number: " +str(count) + parse.printPretty())
         except Exception:
            print ("Error occured at line: " +str(count))
            print ("Error occured at line: " +str(count) +": " + parse.terrainString)
            errors.append(parse.terrainString)
            e.write(parse.terrainString + "\n")

      print ("Total records: " + str(len(labelList)))
      print ("Processing Errors: " + str(len(errors)))

   f.close()
   s.close()
   e.close()
   pass

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
      opts, args = getopt.getopt(argv,"hi:o:s:f:",["ifile=","ofile=","string=", "format="])
   except getopt.GetoptError:
      print (scriptname+ ' -i <inputfile> -o <outputfile>')
      sys.exit(2)

   for opt, arg in opts:
      if opt == '-h':
         print (scriptname +
                  ' -i <inputfile>'
                  ' -o <outputfile>'
                  ' -s <string>')
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputfile = arg

      elif opt in ("-o", "--ofile"):
         successfile = arg
      elif opt in ("-s", "--string"):
         string = arg
      elif opt in ("-f", "--format"):
         format = arg

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
import csv


def printPretty(self):
  q = self._parseProcess()
  prettyProcess = ""
  for key in q.keys():
      prettyProcess += "\n Process String: " + q[key]

  p = []
  for key in self._parseComponents().keys():
    component = self._parseComponents()[key]

    p.append(
      " Raw String: "+ component.string +
      " - Processed String " + repr(component.groupdict())
    )
  return  "\nBase String: " + self.terrainString + "\nComponent Delimiters: " + repr(self._getComponentDelimiters())+ " \nProcesses:" + prettyProcess +"\nComponents\n"+ ("\n").join(p)

def printCSV(self):
  from csvFields import labelFields
  csvArray = []
  for i in labelFields.keys():
    # csvArray.append(labelFields[i])
    csvArray.append("")
  csvArray[labelFields["LABEL_FNL"]] = self.terrainString

  # Hack - if the label field begins with a / assign the symbol to
  # PARTCOV_A
  if self.terrainString.find("/") == 0:
    csvArray[labelFields["PARTCOV_A"]] = self.terrainString[0]

  csvArray = writeComponents(self.components, csvArray)
  csvArray = writeRelationships(self.relationships, csvArray)
  csvArray = writeGeomorphologicalProcess(self.geomorphologicalProcess, csvArray)
  csvArray = writeParsedComponents(self._parseComponents(), csvArray)
  csvArray = writeParsedProcess(self._parseProcess(), csvArray)

  return csvArray

def writeComponents(c, csvArray):
  from csvFields import labelFields

  labels= [
          labelFields["COMP_A"],
          labelFields["COMP_B"],
          labelFields["COMP_C"],
          labelFields["COMP_D"]]

  l = dict(zip(labels, c))
  for i in l.keys():
    csvArray[i] = l[i]

  return csvArray

def writeRelationships(c, csvArray):
  from csvFields import labelFields

  for i in c.keys():
    csvArray[labelFields[i]] = c[i]
  return csvArray

def writeGeomorphologicalProcess(c, csvArray):
  from csvFields import labelFields
  csvArray[labelFields["PROCESS"]] = c
  return csvArray

def writeParsedComponents(c, csvArray):
  from csvFields import labelFields

  componentIter = {
    "componentA": "A",
    "componentB": "B",
    "componentC": "C",
    "componentD": "D"
  }

  for key in c.keys():
    iter = componentIter[key]
    m = c[key].groupdict()

    try:
      csvArray[labelFields["TEXTURE1_"+iter]] = m['texture'][0]
      csvArray[labelFields["TEXTURE2_"+iter]] = m['texture'][1]
      csvArray[labelFields["TEXTURE3_"+iter]] = m['texture'][2]
    except IndexError:
      pass
    try:
      csvArray[labelFields["AGE_"+iter]] = m['age']
    except IndexError:
      pass
    try:
      csvArray[labelFields["EXPRSN1_"+iter]] = m['surfaceExpression'][0]
      csvArray[labelFields["EXPRSN2_"+iter]] = m['surfaceExpression'][1]
      csvArray[labelFields["EXPRSN3_"+iter]] = m['surfaceExpression'][2]
    except IndexError:
      pass
    try:
      if len(m['surficialMaterial']) > 1:
        csvArray[labelFields["QUALIFIER" +iter]] = m['surficialMaterial'][1:]
      csvArray[labelFields["MATERIAL_"+iter]] = m['surficialMaterial']
    except IndexError:
      pass
  return csvArray

def writeParsedProcess(c, csvArray):
  from csvFields import labelFields
  PROCESS_CLASS_LIMIT = 2

  processIter = {
    "processA": "A",
    "processB": "B",
    "processC": "C",
  }

  for key in c.keys():
    iter = processIter[key]
    m = c[key]

    y = 0
    for i in m:
      if y == 0:
         csvArray[labelFields["PRO_QUAL_"+iter]] = m[y]
      elif y <= PROCESS_CLASS_LIMIT:
        csvArray[labelFields["PROCLASS"+str(y)+iter]]=m[y]
      y += 1
    csvArray[labelFields["PROCESS_"+iter]] = m
  return csvArray
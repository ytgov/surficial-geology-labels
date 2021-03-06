
class TerrainClassification:
  """
  Takes a Yukon cerrain classification label and parses it into its constituent
  parts.
  """

  import re

  def __init__(self, terrainString = ""):
    self.terrainString = terrainString
    self.geomorphologicalProcess = self._getGeomorphologicalProcess()
    self.components = self._getComponents()
    self.relationships = self._getComponentDelimiters()
    # self.terrains =

  def printCSV(self):
    import formatters
    return formatters.printCSV(self)

  ### RegExes ###

  def printPretty(self):
    import formatters
    return formatters.printPretty(self)

  # componentRegEx = re.compile("\W")
  componentRegEx = re.compile(r"[\\.///]")

  doubleForwardslashFix = re.compile(r"//")
  singleBackslashFix = re.compile(r"\\")
  processParseRegEx = re.compile(r"""
    #  (?P<process>[A-Z])
    #  (?P<class>[a-z]{,3})
    #  (?P<qualifier>[A?|I?])
     [A-Z]”?[a-z]{,3}A?I?
  """, re.VERBOSE)

  componentParseRegEx = re.compile(r"""

    (?P<texture>^[a-z]{,3})           #TEXTURE - up to 3 lower case letters
                                      #in front of surficial material

    (?P<surficialMaterial>[A-Z]{,2})  # SURFICIAL MATERIAL -
                                      # The first single upper case letter
                                      # shown in map unit. The upper case letter
                                      # immediately following surficial
                                      # material is the glacial or activity
                                      # QUALIFIER.)

    (?P<surfaceExpression>[a-z]{,3})  # SURFACE EXPRESSION - up to 3 lower case
                                      #letters following surficial material

    (?P<age>[A-Z><]{,2})               #AGE - single upper case letter following
                                      #surface expression
    """, re.VERBOSE)



  ### Methods ###
  def _getComponentDelimiters(self):
    string = self.terrainString.split("-")[0]

    #if the label begins with a / don't include it as a delimeter
    #instead the / symbol should by put into the PARTCOV_A CSV field
    if string.find("/") == 0:
      string = string[1:]

    y= {}
    c = []
    for p in self.componentRegEx.finditer(string):
      c.append(p.group())
      # print (self.terrainString[p.end()-2:p.end()+1])
    labels = ["RELATIONAB", "RELATIONBC", "RELATIONCD"]

    for i in range(len(c)):

      if i==0:
        y.update({"RELATIONAB":c[0]})
      if i==1:
        y.update({"RELATIONBC":c[1]})
      if i == 2:
        y.update({"RELATIONCD":c[2]})
    return y

  def _getGeomorphologicalProcess(self):
    """
    Returns the Geomorphological Process for the label.
    GEOMORPHOLOGICAL PROCESS is up to 3 upper case letters following dash “-”.
    Lower case indicate sub classes.
    Per Terrain_Classification_System_summary.pdf
    https://ygsftp.gov.yk.ca/YGSIDS/compilations/Surficial_2014_04_08/Terrain_Classification_System_summary.pdf

    """
    p = self.terrainString.split("-")
    if len(p) > 1:
      return self.terrainString.split("-")[1]
    return ""

  def _getComponents (self):
    """
    The raw components found in terrain string using the regex defined in
    componentRegEx.
    """
    group = self.componentRegEx.split(self.terrainString.split("-")[0])
    return list(filter(None, group))

  def _parseProcess(self):

    process = {
      "processA": "",
      "processB": "",
      "processC": "",
    }
    if not self.geomorphologicalProcess:
      return process
    c = self.processParseRegEx.findall(self.geomorphologicalProcess)
    try:
      if c[0]:
        process.update({"processA":c[0]})
      if c[1]:
        process.update({"processB":c[1]})
      if c[2]:
        process.update({"processC":c[2]})
    except:
      pass
    return process

  def _parseComponents(self):
    """
    Takes list of terrain components.
    \nReturns a list of 're' objects which match the regex defined in componentParseRegEx.
    \nTo see the complete output use m.groupdict()
    \nTo see just values use m.groupdict().values()
    """

    parsedComponents = []
    components = ["componentA", "componentB", "componentC", "componentD"]
    for i in self.components:
      parsedComponents.append(self.componentParseRegEx.search(i))

    l = dict(zip(components, parsedComponents))

    return l

    # return parsedComponents <--- return this if you want printPretty to work
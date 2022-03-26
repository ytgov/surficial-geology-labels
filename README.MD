# Parse Surficial Labels
A python script to parse surficial geology labels for Yukon government's [Terrain Classification System](https://ygsftp.gov.yk.ca/YGSIDS/compilations/Surficial_2014_04_08/Terrain_Classification_System_summary.pdf).

## Instructions
1. Ensure a copy of (Python)[https://www.python.org/downloads/] is installed. *This app was developed using Python 3.10.2*
2. Download the code via the [Github page](https://github.com/ryanjagar/surficial-geology-labels) or using git:

    `git clone https://github.com/ryanjagar/surficial-geology-labels.git`

3. Change into the project directory
4. Run `python parseLabels.py be sure to include any relevant options


## Command Line Options
-h: Display the help message

-i, --ifile:
the name of the file to be parsed

-o, --ofile:
the name of the file where the output will be written

-s, --string:
a single string to be parsed. eg. *"sgFGptM.dsmMbM//xsCv\zcLGpM-RdsbXsVI"*

-f --format:
*Optional* the desired output format; either pretty or csv

## Questions
Questions or comements? Please contact the [Yukon Geological Survey](https://yukon.ca/en/yukon-geological-survey)





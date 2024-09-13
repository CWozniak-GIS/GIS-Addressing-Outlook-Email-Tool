# GIS-Addressing-Outlook-Email-Tool
A Custom ArcGIS Pro Python Geoprocessing Tool for Generating a Template Outlook Email from GIS Address Data

# Generate E911 Address Assignment Notification

### Purpose:
• To quickly generate an authoritative email, which will open as a draft in a new Outlook window, from an input address point selection.
• The tool accommodates many optional parameters relevant to a local government addressing administrator's workflow.
• It is highly recommended for organizations to set up defaults for input feature layers and fields so the tool can be run quickly after the dialogue box opens.

### Note:
There are two geoprocessing tools within the .atbx. The one labeled with "- Bedford County" is identical to the "- template" GP tool,
with the exception of configured default parameters. This should be used as a refrence to configure your own organization's defaults.

### Tool Requirements:
• ArcGIS Pro to open geoprocessing tool in .atbx toolbox
• Microsoft Office / Outlook
• Cloned python environment with 'pywin32' module installed via package manager (read Documentation for more details)

### Author:
Chris Wozniak (Christopher Wozniak)
Academic: cwoznia4@jh.edu | Work: cwozniak@bedfordcountyva.gov | Personal: chris@myyahoo.com

### License:
MIT License

Copyright (c) 2024 Chris Wozniak (Christopher Wozniak)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


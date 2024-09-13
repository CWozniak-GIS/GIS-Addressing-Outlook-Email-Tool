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

# Raw Python Code (Main Script):
    ###
    # Title: Generate E911 Address Assignment Email in Outlook
    # Author: Chris Wozniak
    # Date: 12/18/2023
    ###
    
    import arcpy, sys
    
    try:
        import win32com.client # requires pywin32 module
    except ImportError:
        arcpy.AddError("Missing necessary module. Please add package 'pywin32' to the active environment.")
        sys.exit(1)
    
    def remove_extra_spaces(string):
        words = string.split()
        return " ".join(words)
    
    def remove_empty_fields(fields_list):
        return [fld for fld in fields_list if len(fld.strip()) > 0]
    
    def check_null(arg, meta_string):
        if arg is None or len(arg.strip()) == 0:
            arcpy.AddWarning(f"Possible <NULL> or 'None' value in {meta_string} field. Carefully review input data and generated email.")
            return ''
        else:
            return arg
    
    def emailAppend(email_body, string_to_add):
        email_body += f"\n{string_to_add}\n"
        return email_body
    
    def generateEmail(to, subject, attachment, body):
        try:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = to
            mail.Subject = subject
            if attachment:
                mail.Attachments.Add(attachment)
            mail.Body = body
            mail.Display()
            return arcpy.AddMessage(f"Email (subject: {subject}) successfully generated.")
        except:
            arcpy.AddError("Error contacting Outlook application. Please ensure Outlook is installed and "
                            "that you are signed into Microsoft Office 365 within the same OS environment.")
            return sys.exit(1)
    
            ###
    
    if __name__ == '__main__':
    
        # Defining inputs
    # email inputs
        email_recipient = arcpy.GetParameterAsText(0)
        custom_subject = arcpy.GetParameterAsText(1)
        attachment_path = arcpy.GetParameterAsText(2)
        email_body_inputs = arcpy.GetParameter(21)
    
    # defining administrative inputs
        permit_id = arcpy.GetParameterAsText(3)
        permit_type = arcpy.GetParameterAsText(4)
    
    # address point inputs
        in_addresspoint_fc = arcpy.GetParameterAsText(5)
    
        parsed_input = arcpy.GetParameter(6) #boolean to allow user to input parsed values to build full address string
        address_num = arcpy.GetParameterAsText(7)
        pre_dir = arcpy.GetParameterAsText(8)
        st_name = arcpy.GetParameterAsText(9)
        suffix_dir = arcpy.GetParameterAsText(10)
        st_type = arcpy.GetParameterAsText(11)
        unit = arcpy.GetParameterAsText(12)
    
        parsed_list = [address_num, pre_dir, st_name, st_type, suffix_dir, unit]
        parsed_list = remove_empty_fields(parsed_list)
    
        street_address = arcpy.GetParameterAsText(13) #alternatively, allow full address field input
    
        postal_city = arcpy.GetParameterAsText(14)
        zip_code = arcpy.GetParameterAsText(15)
        state = arcpy.GetParameterAsText(16)
        postal_flds = [postal_city, zip_code, state]
    
    # defining parcel / lot inputs
        in_parcel_fc = arcpy.GetParameterAsText(17)
        parcel_id = arcpy.GetParameterAsText(18)
        lot_num = arcpy.GetParameterAsText(19)
        sublot_num = arcpy.GetParameterAsText(20)
    
            ###
        # Checks and preparatory processes
    
        address_feature_count = int(arcpy.GetCount_management(in_addresspoint_fc)[0])
        if address_feature_count >= 2:
            arcpy.AddError(f"There are {address_feature_count} address point features in the current selection."
                           f" Please select only one address point feature when running this tool.")
            sys.exit(1)
    
        if in_parcel_fc:
            parcel_feature_count = int(arcpy.GetCount_management(in_parcel_fc)[0])
            if parcel_feature_count >= 2:
                arcpy.AddError(f"There are {parcel_feature_count} parcel features in the current selection."
                               f" Please select only one parcel feature when running this tool.")
                sys.exit(1)
    
                ###
            # Extracting Data From Input Feature Class(es)
    
    # Address Point
        address_flds = []
        if parsed_input:
            address_flds = parsed_list + postal_flds
        else:
            address_flds = [street_address] + postal_flds
    
        with (arcpy.da.SearchCursor(in_addresspoint_fc, address_flds) as cursor):
            for row in cursor:
        ## parsed path
                if parsed_input:
                    parsed_list_len = (len(parsed_list))
                    street_address_string = remove_extra_spaces(" ".join(check_null(row[i], "a parsed") for i in range(0, parsed_list_len)))
                    postal_city_string = check_null(row[(0+parsed_list_len)], postal_city)
                    zip_code_string = check_null(str(row[(1+parsed_list_len)]), zip_code)
                    state_string = check_null(row[(2+parsed_list_len)], state)
        ## full path
                else:
                    street_address_string = check_null(row[0], street_address)
                    postal_city_string = check_null(row[1], postal_city)
                    zip_code_string = check_null(str(row[2]), zip_code)
                    state_string = check_null(row[3], state)
    
                whole_address_string = f"{street_address_string.title()}, {postal_city_string.title()}, {state_string.upper()} {zip_code_string}"
    
            # Parcel / Lot
        if in_parcel_fc:
            parcel_flds = [parcel_id, lot_num, sublot_num]
            parcel_flds = remove_empty_fields(parcel_flds)
            with (arcpy.da.SearchCursor(in_parcel_fc, parcel_flds) as cursor):
                for row in cursor:
                    parcel_id_string = check_null(str(row[0]), parcel_id)
                    lot_num_string = check_null(str(row[1]), lot_num)
                    if sublot_num:
                        sublot_num_string = check_null(str(row[2]), sublot_num) # will not always have a value
                        lot_num_string += sublot_num_string
    
        # A warning check to help user catch a mis-selection of a parcel feature
            arcpy.SelectLayerByLocation_management(in_addresspoint_fc, "", in_parcel_fc, "", "SUBSET_SELECTION")
            address_feature_count = int(arcpy.GetCount_management(in_addresspoint_fc)[0])
            if address_feature_count == 0:
                arcpy.AddWarning(f"The selected address point is not within the selected parcel. Carefully review selections and ensure this is "
                            f"intentional. Current selection is: {whole_address_string} and {parcel_id_string}, on lot {lot_num_string}.")
    
            ###
        # Generating email components
        # Feel free to modify structures of email subject and body to your organization's needs
    
        # Email Subject
        if permit_id:
            email_subject = f"Notification of New 911 Address for {permit_type} Bldg. Permit No. {permit_id}"
        elif in_parcel_fc:
            email_subject = f"Notification of New 911 Address for Structure on Lot {lot_num_string}"
        else:
            email_subject = f"Notification of New 911 Address"
    
        if custom_subject:
            email_subject = custom_subject + email_subject
    
        # email body
    
        email_body = "Hello, \n"
    
    # Main E911 address notification statement
        if permit_id:
            if in_parcel_fc:
                email_body = emailAppend(email_body,f"The new E911 address for the {permit_type} Permit ({permit_id}) on Parcel Number "
                                   f"{parcel_id_string} (Lot {lot_num_string}) has been assigned as {whole_address_string}.")
            else:
                email_body = emailAppend(email_body,f"The new E911 address for the {permit_type} Permit ({permit_id}) has been assigned as {whole_address_string}.")
        elif in_parcel_fc:
            email_body = emailAppend(email_body,f"The new E911 address on Parcel Number {parcel_id_string} (Lot {lot_num_string}) has been assigned as {whole_address_string}.")
    
        else:
            email_body = emailAppend(email_body,f"A new E911 address has been assigned as {whole_address_string}.")
            arcpy.AddWarning("No permit or parcel specified. Remember to add context for E911 address assignement reason to email before sending.")
    
    # modular components added to email body - from original email template from Bedford County, VA's GIS Office
        email_body = emailAppend(email_body, ("Reminder: Before final inspection, the structure numbers are required to be posted on the"
        " building near the front door in addition to, if the numbers are not easily visible/legible from the "
        "street, also requiring to be posted at the driveway entrance. See attached guidelines for specific requirements."))
    
        email_body = emailAppend(email_body, (f"When the structure is near completion, the owner will need to contact the {postal_city_string} "
        f"Post Office to have this address established with USPS for local delivery. The USPS website states, “Customers must contact "
        f"their local post office to establish delivery for any new construction prior to submitting any Change of Address Request.”"))
    
    # custom inputs
        if email_body_inputs:
            for input in email_body_inputs:
                email_body = emailAppend(email_body, input)
    
    # signature
        email_body = emailAppend(email_body, "Please let me know if you have any questions.")
    
            ###
        # Generating email in Outlook
        generateEmail(email_recipient, email_subject, attachment_path, email_body)

# Raw Python Code (GP Tool Behavior Validation Script):
    class ToolValidator:
        # Class to add custom behavior and properties to the tool and tool parameters.
    
        def __init__(self):
            # set self.params for use in other function
            self.params = arcpy.GetParameterInfo()
    
        def initializeParameters(self):
            # Customize parameter properties.
            # This gets called when the tool is opened.
            return
    
        def updateParameters(self):
            # Modify parameter values and properties.
            # This gets called each time a parameter is modified, before standard validation.
    
            # Control the parsed fields inputs
            if self.params[6].value == True:
                self.params[13].enabled = False
                for i in range(7, 13):
                    self.params[i].enabled = True
            else:
                self.params[13].enabled = True
                for i in range(7, 13):
                    self.params[i].enabled = False
    
            # Control the Administrative inputs
            if self.params[3].value:
                self.params[4].enabled = True
            else:
                self.params[4].enabled = False
    
            # Control the Parcel / Lot Information inputs
            if self.params[17].value:
                self.params[18].enabled = True
                self.params[19].enabled = True
                self.params[20].enabled = True
            else:
                self.params[18].enabled = False
                self.params[19].enabled = False
                self.params[20].enabled = False
    
            return
    
        def updateMessages(self):
            # Customize messages for the parameters.
            # This gets called after standard validation.
    
            # Control the parsed fields inputs
            if self.params[6].value == True:
                required_parsed = [7, 9, 11]
                #optional_parsed = [8, 10, 12]
                for i in required_parsed:
                    if not self.params[i].valueAsText:
                        self.params[i].setIDMessage('ERROR', 735)
                    else:
                        self.params[i].clearMessage()
            elif self.params[6].value == False and not self.params[13].valueAsText:
                self.params[13].setIDMessage('ERROR', 735)
            else:
                self.params[13].clearMessage()
    
            # Control the Administrative inputs
            if self.params[3].value and not self.params[4].valueAsText:
                self.params[4].setIDMessage('ERROR', 735)
            else:
                self.params[4].clearMessage()
    
            # Control the Parcel / Lot Information inputs
            if self.params[17].value:
                for i in range(18, 20):
                    if not self.params[i].valueAsText:
                        self.params[i].setIDMessage('ERROR', 735)
                    else:
                        self.params[i].clearMessage()
    
            return
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

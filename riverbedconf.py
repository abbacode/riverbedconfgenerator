__version__ = 1.0

import xlrd
import re

class Config(object):
    def __init__(self):
        self.raw_database = {}      # Raw information read from database.xls
        self.database = {}          # Real database based on raw_database, but has had invalid entries removed
        self.config_templates = {}  # Stores configuration templates

        self.read_database_from_file()     # Read the content from database.xls
        self.get_templates_from_database() # Read 'templates' from raw_database and place into config_templates
        self.get_content_from_database()   # Read 'config' from raw_database and place into database
        self.generate_all_configuration()  # Generate required configuration output to .txt files


    #----------------------------------------------------------------
    # Read information from the database.xlsx
    # Information will be stored by worksheet name, row, column name
    #-----------------------------------------------------------------
    def read_database_from_file(self):
        try:
            wb = xlrd.open_workbook('database.xls')
        except:
            print ("Cannot open database.xls file, aborting script")
            exit()

        temp_db = []

        for i, worksheet in enumerate(wb.sheets()):
            header_cells = worksheet.row(0)
            num_rows = worksheet.nrows - 1
            curr_row = 0
            header = [each.value for each in header_cells]
            while curr_row < num_rows:
                curr_row += 1
                row = [int(each.value) if isinstance(each.value, float)
                       else each.value
                       for each in worksheet.row(curr_row)]
                value_dict = dict(zip(header, row))
                temp_db.append(value_dict)
            else:
                self.raw_database[worksheet.name] = temp_db
                temp_db = [] 


    # -----------------------------------------------------------------------------------------
    # Each row must have a value defined for the following columns to be considered valid
    # -----------------------------------------------------------------------------------------
    def valid_entry_in_database(self,database_entry):
        VALID_FIELDS = ['HOSTNAME', 'TEMPLATE']

        for field in VALID_FIELDS:
            if not database_entry[field]:
                return False
        return True

    # -----------------------------------------------------------------------------------------
    # Iterate over all entries in the raw_database and extract valid entries into database
    # -----------------------------------------------------------------------------------------
    def get_content_from_database(self):
        for row in self.raw_database["variables"]:
            if self.valid_entry_in_database(row):
                host = str(row["HOSTNAME"]).strip()
                if host not in self.database:
                    tempObj = Riverbed()
                    tempObj.hostname            = str(row["HOSTNAME"]).strip()
                    tempObj.template            = str(row["TEMPLATE"]).strip()
                    tempObj.inpath0_0_ip        = str(row["INPATH0_0_IP"]).strip()
                    tempObj.inpath0_0_subnet    = str(row["INPATH0_0_SUBNET_MASK"]).strip()
                    tempObj.inpath0_0_default   = str(row["INPATH0_0_DEFAULT_GATEWAY"]).strip()
                    tempObj.inpath0_0_vlan      = str(row["INPATH0_0_VLAN"]).strip()
                    tempObj.inpath0_1_ip        = str(row["INPATH0_1_IP"]).strip()
                    tempObj.inpath0_1_subnet    = str(row["INPATH0_1_SUBNET_MASK"]).strip()
                    tempObj.inpath0_1_default   = str(row["INPATH0_1_DEFAULT_GATEWAY"]).strip()
                    tempObj.inpath0_1_vlan      = str(row["INPATH0_1_VLAN"]).strip()
                    tempObj.primary_ip          = str(row["PRIMARY_IP"]).strip()
                    tempObj.primary_subnet      = str(row["PRIMARY_SUBNET"]).strip()
                    tempObj.primary_default     = str(row["PRIMARY_DEFAULT"]).strip()
                    tempObj.dns                 = str(row["DNS"]).strip()
                    tempObj.ntp                 = str(row["NTP"]).strip()

                    # A valid row was detected, create a new row in the final database
                    self.database[host] = tempObj

    # --------------------------------------------------------------------------------------------------
    # Iterate over all template entries in raw_database and extract valid entries into config_templates
    # --------------------------------------------------------------------------------------------------
    def get_templates_from_database(self):
        temp_list = []
        for row in range(len(self.raw_database["config-templates"])):
            row_value = self.raw_database["config-templates"][row]["Enter config templates below this line:"]
            match_new_template = re.search(r'Config Template: \[(.*?)\]', row_value,re.IGNORECASE)
            if not row_value:
                continue
            if match_new_template:
                new_template_name = match_new_template.group(1)
                new_template_config = []
                self.config_templates[new_template_name] = new_template_config
                continue
            self.config_templates[new_template_name].append(row_value)

    #---------------------------------------------------------------------------------------------------------
    # Read through the template used for each host and replacement dynamic variable values with correct value
    #---------------------------------------------------------------------------------------------------------
    def prepare_configuration(self,host):

        host.config = self.config_templates[host.template]
        new_config = []

        for line in host.config:
            DYNAMIC_VARIABLES = re.findall(r'\[(.*?)\]', line)
            if (DYNAMIC_VARIABLES):
                # could be more than one variable on each line, iterate over them
                for variable in DYNAMIC_VARIABLES:
                    line = line.replace("[","")
                    line = line.replace("]","")
                    new_variable_value = self.get_variable_value(host,variable)
                    if new_variable_value:
                        line = line.replace(variable,new_variable_value)
                    else:
                        line = line.replace(variable,"[ERROR DYNAMIC VARIABLE VALUE NOT FOUND]")
            new_config.append(line)
        host.config = new_config

    #-------------------------------------------------
    # Search for an exact string, not a partial match
    #-------------------------------------------------
    def find_exact_string(self,word):
        return re.compile(r'^\b({0})$\b'.format(word),flags=re.IGNORECASE).search

    #-------------------------------------------------------------------------
    # Iterate over the database and pull the record that matches the hostname
    #-------------------------------------------------------------------------
    def get_host(self,hostname):
        for host in self.database:
            if self.find_exact_string(hostname)(host):
                return self.database[host]
        return None

    def get_variable_value(self,host,dynamic_variable):
        for keyword, value in vars(host).items():
            if dynamic_variable in keyword.upper():
                return (value)
        return None



    #---------------------------------------------
    # Generate configuration for a single entry
    #---------------------------------------------
    def generate_configuration(self,hostname):
        host = self.get_host(hostname)
        if not host:
            return
        self.prepare_configuration(host)

        with open(hostname+".txt","w") as output_file:
            print ("=============================",file=output_file)
            print (" Values for host: '{}'      ".format(host.hostname),file=output_file)
            print ("=============================",file=output_file)
            print ("Template           : {}".format(host.template),file=output_file)
            print ("Primary IP         : {}".format(host.primary_ip),file=output_file)
            print ("Primary Subnet     : {}".format(host.primary_subnet),file=output_file)
            print ("Primary Default    : {}".format(host.primary_default),file=output_file)
            print ("NTP                : {}".format(host.ntp),file=output_file)
            print ("DNS                : {}".format(host.dns),file=output_file)
            print ("Inpath0_0 IP       : {}".format(host.inpath0_0_ip),file=output_file)
            print ("Inpath0_0 Subnet   : {}".format(host.inpath0_0_subnet),file=output_file)
            print ("Inpath0_0 Default  : {}".format(host.inpath0_0_default),file=output_file)
            print ("Inpath0_0 VLAN     : {}".format(host.inpath0_0_vlan),file=output_file)
            print ("Inpath0_1 IP       : {}".format(host.inpath0_1_ip),file=output_file)
            print ("Inpath0_1 Subnet   : {}".format(host.inpath0_1_subnet),file=output_file)
            print ("Inpath0_1 Default  : {}".format(host.inpath0_1_default),file=output_file)
            print ("Inpath0_1 VLAN     : {}".format(host.inpath0_1_vlan),file=output_file)
            print ("=============================",file=output_file)
            print (" Config template: '{}'       ".format(host.template),file=output_file)
            print ("=============================",file=output_file)
            for line in host.config:
                print (line,file=output_file)

    #-------------------------------------------------------------
    # Iterate over all entries and generate a config for each one
    #-------------------------------------------------------------
    def generate_all_configuration(self):
        print ("Generating configuration files.....")
        for hostname in sorted(self.database):
            self.generate_configuration(hostname)
            print ("  ++ {}.txt generated".format(hostname))

    def __repr__(self):
        return repr(self.database)

#---------------------------------------------------------------------
# Update the template structure if you need to define default values
#---------------------------------------------------------------------
class Riverbed(object):
    def __init__(self):
        self.hostname = "TBD"
        pass
    def __repr__(self):
        return self.hostname

# Run the script
config = Config()
print (config)
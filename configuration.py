import configparser


class Config:

    def __init__(self):
        config = configparser.ConfigParser()

        config.read('config\config.ini')

        self.source = config['SOURCES']['SOURCE']
        self.report = config['SOURCES']['REPORT']
        self.path_gen = config['SOURCES']['GEN']
        self.send_to = config['RECEIVERS']['SEND_TO']

        self.ini_cell = config['TEMPLATE']['INI_CELL']

        # Workbook configuration
        self.config_account_col = config['CONFIG_WB']['config_account_col']
        self.config_name_col = config['CONFIG_WB']['config_name_col']
        self.config_instance_col = config['CONFIG_WB']['config_instance_col']
        self.config_user_col = config['CONFIG_WB']['config_user_col']
        self.config_password_col = config['CONFIG_WB']['config_password_col']
        self.config_server_col = config['CONFIG_WB']['config_server_col']
        self.config_sysnr_col = config['CONFIG_WB']['config_sysnr_col']
        self.config_client_col = config['CONFIG_WB']['config_client_col']
        self.config_saprouter_col = config['CONFIG_WB']['config_saprouter_col']

        self.config_systemid_col = config['CONFIG_WB']['config_systemid_col']
        self.config_system_col = config['CONFIG_WB']['config_system_col']
        self.config_enviroment_col = config['CONFIG_WB']['config_enviroment_col']
        self.config_type_col = config['CONFIG_WB']['config_type_col']

        self.config_version_portal_col = config['CONFIG_WB']['config_version_portal_col']

        self.sheet_connections = config['CONFIG_WB']['sheet_connections']
        self.sheet_dashboard = config['CONFIG_WB']['sheet_dashboard']

        self.conf_threshold_first = int(config['THRESHOLD']['first'])
        self.conf_threshold_second = int(config['THRESHOLD']['second'])
        self.conf_threshold_third = int(config['THRESHOLD']['third'])

        self.subject = config['GLOBAL']['subject']
        self.send_email = bool(config['GLOBAL']['send_email'])

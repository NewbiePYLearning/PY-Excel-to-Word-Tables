data migration

    Through discussion with the customer, the MOP document has been solidified into a template (including the style and font).
    
    These settings should be kept in the new MOP document. The migration data is designed by delivery engineers and stored in Excel files in a fixed format.
    
    The migration data varies according to the sites. Therefore, generate different MOP documents for different sites.
    
    :param input1_MOP_template: Sample MOP Doc（SiteName：Den Haag）
    
    :param input2_MOP_basic_info: MOP basic info table
    
    :param input_dir_Site_Port_Matrix_Tables: The file contain port matrix of the sites, they are Site1_Port_Matrix_Table.xlsx和Site2_Port_Matrix_Table.xlsx
    
    :param output_folder_of_MOPs: The specified output folder, need creat the folder if not exist：
    
        Generate the MOP doc for site “Site1”和“Site2”. The file is named as the format "MOP_<siteName>_V1.0.docx"
        
    :return:

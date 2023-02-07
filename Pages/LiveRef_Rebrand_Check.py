import time
import openpyxl
from re import search
from selenium.webdriver.support.wait import WebDriverWait
from Pages.Base import Base
from Pages.LiveRef_SearchPublicationsPage import SearchPublicationsPage
from Pages.SLRReportPage import SLRReport
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support.ui import Select


class LiveRef_Rebrand(Base):
    
    """Constructor of the LiveRef_Rebrand class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)
    
    def validate_liveref_rebrand(self, TestData, env):

        user_role = self.get_text("user_role", env)
        if search("LiveRef", user_role) and not search("PubTracker", user_role):
            self.LogScreenshot.fLogScreenshot(message=f"User Role is : {user_role}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in User role section. Actual value is : {user_role}")
        
        title_bar_txt = self.get_text("title_bar", env)
        if search("LiveRef", title_bar_txt) and not search("PubTracker", title_bar_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Title bar : {title_bar_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in LiveRef Title bar. Actual value is : {title_bar_txt}")

        about_liveref_txt = self.get_text("about_live_ref", env)
        if search("LiveRef", about_liveref_txt) and not search("PubTracker", about_liveref_txt):
            self.LogScreenshot.fLogScreenshot(message=f"About LiveRef : {about_liveref_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in About LiveRef section. Actual value is : {about_liveref_txt}")
        
        liveref_insights_txt = self.get_text("dashboard_insights", env)
        if search("LiveRef", liveref_insights_txt) and not search("PubTracker", liveref_insights_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Dashboard Insights : {liveref_insights_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Dashboard Insights section. Actual value is : {liveref_insights_txt}")
        
        liveref_updates_txt = self.get_text("dashboard_updates", env)
        if search("Updates", liveref_updates_txt) and not search("Congress News", liveref_updates_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Dashboard Updates : {liveref_updates_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Dashboard Updates section. Actual value is : {liveref_updates_txt}")
        
        self.go_to_page("searchpublications_button", env)
        self.click("sourceofdata_section", env)
        src_updates_txt = self.get_text("sourceofdata_allupdates_btn", env)
        if search("All Updates", src_updates_txt) and not search("All Congressess/Libraries", src_updates_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Source of Data section Updates text : "
                                                      f"{src_updates_txt}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Source of Data -> All updates radio "
                            f"button section. Actual value is : {src_updates_txt}")

        src_param_header = self.get_text("selected_src_param_head", env)
        if search("Source of data", src_param_header) and not search("Congressess", src_param_header):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Source of Data selected parameter "
                                                      f"label text : {src_param_header}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Source of Data -> Selected parameter section. "
                            f"Actual value is : {src_param_header}")

        src_param_body = self.get_text("selected_src_param_body", env)
        if search("All Updates", src_param_body) and not search("All Congressess", src_param_body):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Source of Data selected parameter "
                                                      f"value text : {src_param_body}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Source of Data -> Selected parameter section. "
                            f"Actual value is : {src_param_body}")
            
        self.click("cat_of_evidence_section", env)
        cat_evidence_txt = self.get_text("cat_of_evidence_section", env)
        if search("Select Category of Evidence", cat_evidence_txt) and \
                not search("Select a Type of Study/GVD Chapter", cat_evidence_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Category of Evidence text : "
                                                      f"{cat_evidence_txt}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Category of Evidence section. Actual "
                            f"value is : {cat_evidence_txt}")

        cat_evi_placeholder_txt = self.select_element("cat_of_evidence_placeholder", env).get_attribute("placeholder")
        if search("Search Category of Evidence", cat_evi_placeholder_txt) and \
                not search("Search type of study", cat_evi_placeholder_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Category of Evidence placeholder "
                                                      f"text : {cat_evi_placeholder_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Category of Evidence placeholder text section. "
                            f"Actual value is : {cat_evi_placeholder_txt}")

        cat_evi_param_header = self.get_text("selected_cat_param_head", env)
        if search("Category of Evidence", cat_evi_param_header) and \
                not search("Type of Studies/GVD Chapters", cat_evi_param_header):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Category of Evidence selected "
                                                      f"parameter label text : {cat_evi_param_header}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Category of Evidence -> Selected parameter "
                            f"section. Actual value is : {cat_evi_param_header}")
            
        cat_evi_param_body = self.get_text("selected_cat_param_body", env)
        if search("All Categories of Evidence", cat_evi_param_body) and \
                not search("All Type of Studies", cat_evi_param_body):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Category of Evidence selected parameter "
                                                      f"value text : {cat_evi_param_body}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Category of Evidence -> Selected parameter "
                            f"section. Actual value is : {cat_evi_param_body}")

        self.click("reported_cols_sections", env, UnivWaitFor=10)
        time.sleep(2)
        rpt_col_src_data = self.get_text("reported_col_sourcedata", env)
        if search("Source of Data", rpt_col_src_data) and not search("Conference", rpt_col_src_data):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column1 text : "
                                                      f"{rpt_col_src_data}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column1. Actual value is : "
                            f"{rpt_col_src_data}")

        rpt_col_short_ref = self.get_text("reported_col_shortref", env)
        if search("Short Reference", rpt_col_short_ref):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column2 text : "
                                                      f"{rpt_col_short_ref}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column2. Actual value is : "
                            f"{rpt_col_short_ref}")

        rpt_col_qual_sum = self.get_text("reported_col_qualsummary", env)
        if search("Qualitative Summary", rpt_col_qual_sum) and not search("Main Message", rpt_col_qual_sum):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column3 text : "
                                                      f"{rpt_col_qual_sum}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column3. Actual value is : "
                            f"{rpt_col_qual_sum}")

        rpt_col_evi_cat = self.get_text("reported_col_evicat", env)
        if search("Evidence Category", rpt_col_evi_cat) and not search("Study Time", rpt_col_evi_cat):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column4 text : "
                                                      f"{rpt_col_evi_cat}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column4. Actual value is : "
                            f"{rpt_col_evi_cat}")

        rpt_col_subcat = self.get_text("reported_col_subcat", env)
        if search("Study Type / Subcategory", rpt_col_subcat) and not search("Study Sub-Type", rpt_col_subcat):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column5 text : "
                                                      f"{rpt_col_subcat}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column5. Actual value is : "
                            f"{rpt_col_subcat}")

        rpt_col_pubdetails = self.get_text("reported_col_pubdetails", env)
        if search("Publication Details", rpt_col_pubdetails) and \
                not search("Day,Time and Location", rpt_col_pubdetails):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column6 text : "
                                                      f"{rpt_col_pubdetails}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column6. Actual value is : "
                            f"{rpt_col_pubdetails}")

        rpt_col_reportedvar = self.get_text("reported_col_reportedvar", env)
        if search("Methods / Reported Variables", rpt_col_reportedvar) and \
                not search("Reported Data Variables", rpt_col_reportedvar):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column7 text : "
                                                      f"{rpt_col_reportedvar}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column7. Actual value is : "
                            f"{rpt_col_reportedvar}")

        rpt_col_additional_details = self.get_text("reported_col_additionaldetails", env)
        if search("Additional Details", rpt_col_additional_details) and \
                not search("Others", rpt_col_additional_details):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column8 text : "
                                                      f"{rpt_col_additional_details}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column8. Actual value is : "
                            f"{rpt_col_additional_details}")
            
        rpt_col_quant_results = self.get_text("reported_col_quantresults", env)
        if search("Quantitative Results", rpt_col_quant_results) and not search("Main Results", rpt_col_quant_results):
            self.LogScreenshot.fLogScreenshot(message=f"Search Publications -> Select Reported Column9 text : "
                                                      f"{rpt_col_quant_results}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Search Publications -> Select Reported Column9. Actual value is : "
                            f"{rpt_col_quant_results}")

        self.click("searchpublications_reset_filter", env)
        excel_name = self.srchpub.validate_downloaded_filename('scenario1', TestData, env)
        excel = openpyxl.load_workbook(f'ActualOutputs//{excel_name}')
        excel_sheet = excel.active
        if search("LiveRef Report", excel_sheet['A1'].value) and search("LiveRef Report", excel_sheet.title):
            self.LogScreenshot.fLogScreenshot(message=f"Excel Report : {excel_sheet['A1'].value}",
                                              pass_=True, log=True, screenshot=True)
            self.LogScreenshot.fLogScreenshot(message=f"Excel Report sheet name : {excel_sheet.title}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception("Mismatch found in Downloaded LiveRef report")

        self.go_to_page("managesourcedata_button", env)
        src_data_page_link_text = self.get_text("managesourcedata_button", env)
        if search("Manage Source of data", src_data_page_link_text) and \
                not search("Manage Congresses/Libraries", src_data_page_link_text):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Page Link : {src_data_page_link_text}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data page link. Actual value is : "
                            f"{src_data_page_link_text}")

        src_data_page_name = self.get_text("managesourcedata_pagename", env)
        if search("Manage Source of data", src_data_page_name) and \
                not search("Manage Congresses/Libraries", src_data_page_name):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Page Name : {src_data_page_name}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data page name. Actual value is : "
                            f"{src_data_page_name}")

        add_src_data_btn_txt = self.get_text("add_sourcedata_btn", env)
        if search("Add Source of data", add_src_data_btn_txt) and not search("Add Congress", add_src_data_btn_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Add button : {add_src_data_btn_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Add button. Actual value is : "
                            f"{add_src_data_btn_txt}")

        self.click("add_sourcedata_btn", env)                              
        add_src_data_form_label = self.get_text("mngsrc_form_label", env)
        if search("Add Source of data", add_src_data_form_label) and \
                not search("Add Congress", add_src_data_form_label):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form label : "
                                                      f"{add_src_data_form_label}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form label. Actual value is : "
                            f"{add_src_data_form_label}")
            
        src_form_name = self.get_text("mngsrc_form_name", env)
        if search("Source name", src_form_name) and not search("Congress", src_form_name):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form Name : {src_form_name}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form Name. Actual value is : {src_form_name}")

        src_form_abbreviation = self.get_text("mngsrc_form_abbreviation", env)
        if search("Source Abbreviation", src_form_abbreviation) and \
                not search("Congress short code", src_form_abbreviation):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form Abbreviation : "
                                                      f"{src_form_abbreviation}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form Abbreviation. Actual value is : "
                            f"{src_form_abbreviation}")

        src_form_short_code = self.get_text("mngsrc_form_short_code", env)
        if search("Source short code", src_form_short_code) and not search("Conference", src_form_short_code):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form Short code : "
                                                      f"{src_form_short_code}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form Short code. Actual value is : "
                            f"{src_form_short_code}")

        src_form_startdate = self.get_text("mngsrc_form_startdate", env)
        if search("Source start date", src_form_startdate) and not search("Congress", src_form_startdate):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form Start date : "
                                                      f"{src_form_startdate}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form Start date. Actual value is : "
                            f"{src_form_startdate}")

        src_form_enddate = self.get_text("mngsrc_form_enddate", env)
        if search("Source end date", src_form_enddate) and not search("Congress", src_form_enddate):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form End date : {src_form_enddate}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form End date. Actual value is : "
                            f"{src_form_enddate}")
            
        src_form_logo = self.get_text("mngsrc_form_logo", env)
        if search("Source logo", src_form_logo) and not search("Congress", src_form_logo):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Form Logo : {src_form_logo}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Form Logo. Actual value is : {src_form_logo}")

        src_data_table_heading = self.get_text("mngsrc_table_heading", env)
        if search("Sources", src_data_table_heading) and not search("Congresses", src_data_table_heading):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data -> Table Heading with count : "
                                                      f"{src_data_table_heading}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Table Heading. Actual value is : "
                            f"{src_data_table_heading}")

        src_table_logo = self.get_text("mngsrc_table_sourcelogo_col", env)
        if search("Source logo", src_table_logo) and not search("Congress", src_table_logo):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data table -> Table Logo : {src_table_logo}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Table Logo. Actual value is : "
                            f"{src_table_logo}")

        src_table_name = self.get_text("mngsrc_table_sourcename_col", env)
        if search("Source name", src_table_name) and not search("Congress", src_table_name):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data table -> Table Name : {src_table_name}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Table Name. Actual value is : "
                            f"{src_table_name}")

        src_table_shortcode = self.get_text("mngsrc_table_sourceshrtcode_col", env)
        if search("Source short code", src_table_shortcode) and not search("Congress", src_table_shortcode):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of data table -> Table short code : "
                                                      f"{src_table_shortcode}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Source of data -> Table short code. Actual value is : "
                            f"{src_table_shortcode}")

        self.go_to_page("liveref_managecatevidence_button", env)
        cat_evidence_page_link_txt = self.get_text("liveref_managecatevidence_button", env)
        if search("Manage Category of", cat_evidence_page_link_txt) \
                and not search("Type of Studies/GVD Chapters", cat_evidence_page_link_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Category of Evidence -> Page Link : "
                                                      f"{cat_evidence_page_link_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Category of Evidence -> Page link. Actual value is : "
                            f"{cat_evidence_page_link_txt}")

        cat_evidence_page_name = self.get_text("liveref_managecatevidence_pagename", env)
        if search("Manage Category of", cat_evidence_page_name) and \
                not search("Type of Studies/GVD Chapters", cat_evidence_page_name):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Category of Evidence -> Page Name : "
                                                      f"{cat_evidence_page_name}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Category of Evidence -> Page Name. Actual value is : "
                            f"{cat_evidence_page_name}")
        
        add_cat_evidence_btn_txt = self.get_text("add_catevidence_button", env)
        if search("Add Category of Evidence", add_cat_evidence_btn_txt) and \
                not search("Type of Study", add_cat_evidence_btn_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Category of Evidence -> Add button text : "
                                                      f"{add_cat_evidence_btn_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Category of Evidence -> Add button text. Actual value is : "
                            f"{add_cat_evidence_btn_txt}")

        self.click("add_catevidence_button", env)                              
        add_cat_evidence_form_label = self.get_text("mngcatevidence_form_label", env)
        if search("Add Category of Evidence", add_cat_evidence_form_label) and \
                not search("Type of Study", add_cat_evidence_form_label):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Category of Evidence -> Form label : "
                                                      f"{add_cat_evidence_form_label}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Category of Evidence -> Form label. Actual value is : "
                            f"{add_cat_evidence_form_label}")

        add_cat_evidenceform_name = self.get_text("mngcatevidence_form_name", env)
        if search("Category of Evidence name", add_cat_evidenceform_name) and \
                not search("Type of study", add_cat_evidenceform_name):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Category of Evidence -> Form name : "
                                                      f"{add_cat_evidenceform_name}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Category of Evidence -> Form name. Actual value is : "
                            f"{add_cat_evidenceform_name}")

        cat_evidence_table_heading = self.get_text("mngcatevidence_table_heading", env)
        if search("Category of Evidence", cat_evidence_table_heading) and \
                not search("Type of studies", cat_evidence_table_heading):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Category of Evidence -> table heading : "
                                                      f"{cat_evidence_table_heading}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Manage Category of Evidence -> Table heading. Actual value is : "
                            f"{cat_evidence_table_heading}")
            
        self.go_to_page("liveref_importpublications_button", env)
        import_pub_page_name = self.get_text("liveref_importpublications_pagename", env)
        if search("Import Publications", import_pub_page_name) and \
                not search("Congress", import_pub_page_name) and not search("PubTracker", import_pub_page_name):
            self.LogScreenshot.fLogScreenshot(message=f"Import Publications -> Page Name : {import_pub_page_name}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Import Publications - Page Name. Actual value is : "
                            f"{import_pub_page_name}")
            
        import_pub_download_label = self.get_text("importpublications_fieldtext_download", env)
        if search("Select Source", import_pub_download_label) and not search("Congress", import_pub_download_label):
            self.LogScreenshot.fLogScreenshot(message=f"Import Publications -> Download template Field Label : "
                                                      f"{import_pub_download_label}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Import Publications - Download template Field Label. Actual value is : "
                            f"{import_pub_download_label}")

        ele1 = self.select_element("importpub_select_dropdown", env)
        select_dropdown1 = Select(ele1)
        dropdown_text1 = select_dropdown1.first_selected_option.text   
        if search("Select Source", dropdown_text1) and not search("Congress", dropdown_text1):   
            self.LogScreenshot.fLogScreenshot(message=f"Import Publications -> Download template dropdown text : "
                                                      f"{dropdown_text1}", pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Import Publications - Download template dropdown text. "
                            f"Actual value is : {dropdown_text1}")

        import_pub_upload_label = self.get_text("importpublications_fieldtext_upload", env)
        if search("Select Source", import_pub_download_label) and not search("Congress", import_pub_upload_label):
            self.LogScreenshot.fLogScreenshot(message=f"Import Publications -> Upload template Field Label : "
                                                      f"{import_pub_upload_label}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Import Publications - Upload template Field Label. Actual value is : "
                            f"{import_pub_upload_label}")

        ele2 = self.select_element("importpub_upload_dropdown", env)
        select_dropdown2 = Select(ele2)
        dropdown_text2 = select_dropdown2.first_selected_option.text  
        if search("Select Source", dropdown_text2) and not search("Congress", dropdown_text2):    
            self.LogScreenshot.fLogScreenshot(message=f"Import Publications -> Upload template dropdown text : "
                                                      f"{dropdown_text2}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Import Publications - Upload template dropdown text. "
                            f"Actual value is : {dropdown_text2}")

        import_pub_bullet_txt = self.get_text("importpub_bullet1_text", env)
        if search("Source", import_pub_bullet_txt) and not search("Congress", import_pub_bullet_txt):
            self.LogScreenshot.fLogScreenshot(message=f"Import Publications -> Bullet point : {import_pub_bullet_txt}",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Mismatch found in Import Publications -> Bullet point. Actual value is : "
                            f"{import_pub_bullet_txt}")

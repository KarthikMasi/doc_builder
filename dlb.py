#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pycurl, cStringIO, json

def fetch_json_of_database():
    """
    returns json object containing all records as key value pairs
    :return: list<dict>
    """
    buf = cStringIO.StringIO()
    content = {
                'token': '0D806AD514A962B4C4B1B5F2A5E284A0',
                'content': 'record',
                'format': 'json',
                'type': 'flat',
                'records[0]': 'DL5006',
                'fields[0]': 'dl_age',
                'fields[1]': 'dl_ctopp_el_ss',
                'fields[2]': 'dl_ctopp_nr_ss',
                'fields[3]': 'dl_date',
                'fields[4]': 'dl_doa_45c',
                'fields[5]': 'dl_dob',
                'fields[6]': 'dl_first_name',
                'fields[7]': 'dl_sex',
                'fields[8]': 'dl_towk_exp_ss',
                'fields[9]': 'dl_towk_rec_ss',
                'fields[10]': 'dl_towre_pde_pile_grade',
                'fields[11]': 'dl_towre_pde_ss_grade',
                'fields[12]': 'dl_towre_swe_pile_grade',
                'fields[13]': 'dl_towre_swe_ss_grade',
                'fields[14]': 'dl_towre_twe_pile_grade',
                'fields[15]': 'dl_towre_twe_ss_grade',
                'fields[16]': 'dl_wiscds_ss',
                'fields[17]': 'dl_wj_lwid_pile_grade',
                'fields[18]': 'dl_wj_lwid_ss_grade',
                'fields[19]': 'dl_wj_oc_pile_grade',
                'fields[20]': 'dl_wj_oc_ss_grade',
                'fields[21]': 'dl_wj_pc_pile_grade',
                'fields[22]': 'dl_wj_pc_ss_grade',
                'fields[23]': 'dl_wj_wa_pile_grade',
                'fields[24]': 'dl_wj_wa_ss_grade',
                'forms[0]': 'ctopp2',
                'forms[1]': 'demographics',
                'forms[2]': 'towk',
                'forms[3]': 'towre2',
                'forms[4]': 'wisc_v',
                'forms[5]': 'wjiv',
                'rawOrLabel': 'raw',
                'rawOrLabelHeaders': 'raw',
                'exportCheckboxLabel': 'false',
                'exportSurveyFields': 'false',
                'exportDataAccessGroups': 'false',
                'returnFormat': 'json'
        }
    ch = pycurl.Curl()
    ch.setopt(ch.URL, 'https://redcap.vanderbilt.edu/api/')
    ch.setopt(ch.HTTPPOST, content.items())
    ch.setopt(ch.WRITEFUNCTION, buf.write)
    ch.perform()
    ch.close()
    data = json.loads(buf.getvalue())
    buf.close()
    return data 

#! /usr/local/bin/python

'''
Created on Oct 17, 2016

@author: mike
'''

import xlrd
import time
import datetime



if __name__ == '__main__':
    
    MY_DIR = "/Users/mike/Documents/coc/"
    XL_FILE = "CerebralChaosStats.xlsx"
    XL_FILE_PATH = MY_DIR + XL_FILE
    XL_PLAYER_WS = "Player Metrics"
    XL_WAR_WS = "War Metrics"
    
    workbook = xlrd.open_workbook(XL_FILE_PATH)
    player_worksheet = workbook.sheet_by_name(XL_PLAYER_WS)
    war_worksheet = workbook.sheet_by_name(XL_WAR_WS)
    
    
    #print "Rank:",  worksheet.cell(2, 0).value
    #print "Level:",  worksheet.cell(2, 1).value
    #print "Player:",  worksheet.cell(2, 2).value
    #print "Defensibility:",  worksheet.cell(2, 3).value
    #print "Included:",  worksheet.cell(2, 4).value
    #print "Attack Utilization:",  worksheet.cell(2, 5).value
    #print "Star Differential:",  worksheet.cell(2, 6).value
    #print "Star Differential Average:",  worksheet.cell(2, 7).value
    #print "Total Destruction:",  worksheet.cell(2, 8).value
    
    html_str = ""
    star_diff_color = ""
    star_diff_avg_color = ""
    
    for row_index in range (2,28):
        html_str = html_str + "<tr>\n"
        rank =  player_worksheet.cell(row_index, 0).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(rank) + "</td>\n"
        level =  player_worksheet.cell(row_index, 1).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(int(level)) + "</td>\n"
        player =  player_worksheet.cell(row_index, 2).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(player) + "</td>\n"
        defense =  player_worksheet.cell(row_index, 3).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(round(defense,1)) + "</td>\n"
        included = player_worksheet.cell(row_index, 4).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(int(included)) + "</td>\n"
        attack_util = player_worksheet.cell(row_index, 5).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(round(attack_util,2)*100) + "%</td>\n"
        star_diff = player_worksheet.cell(row_index, 6).value
        if star_diff < 0:
            star_diff_color = "style=\"text-align: center; color: red;\""
        else:
            star_diff_color = "style=\"text-align: center;\""
        html_str = html_str + "  <td " + star_diff_color +" >" + str(int(star_diff)) + "</td>\n"
        star_diff_avg = player_worksheet.cell(row_index, 7).value
        html_str = html_str + "  <td " + star_diff_color +" >" + str(round(star_diff_avg, 1)) + "</td>\n"
        tot_dest = player_worksheet.cell(row_index, 8).value
        html_str = html_str + "  <td style=\"text-align: center;\">" + str(int(tot_dest)) + "</td>\n"
        html_str = html_str + "</tr>\n"
        html_str = html_str + "\n\n\n\n"
            
    #print html_str
    
    
    utc_ts = datetime.datetime.utcnow()
    utc_string = utc_ts.strftime('%Y-%m-%d-%H%M%SZ')
    
    # Open a file
    filename = MY_DIR + "coc-stats-" + utc_string + ".txt"
    print filename
    fo = open(filename, "wb")
    fo.write(html_str);

    # Close opend file
    fo.close()
    
    

            
            

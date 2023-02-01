## Spindles detection on specific sleep stages


# Step 1: import necessary libraries
import yasa
import os
import numpy as np
import pandas as pd
import mne
from pathlib import Path
import logging

#set up logging
logging.basicConfig(filename='logFile.log', level=logging.DEBUG, force=True)
logging.captureWarnings(True)
logging.info("Spindle Density Log File (errors and warnings)")


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#step 2: create output excel file
#name of output excel file to which all the spindle data will be written
outputFile = 'SpindleStats.xlsx'



#if 'SpindleStats.xlsx' does not exist in current folder, then create it
#we need a title sheet because otherwise when use 'append' mode, the
#excel writer complains the sheet doesn't already exist
if not os.path.exists(outputFile):
    with pd.ExcelWriter(outputFile, engine="openpyxl", mode='w') as writer:
        df = pd.DataFrame({'Info': ['Author: Sleep Lab', 'Spindle Density Tables']})
        df.to_excel(writer, sheet_name='Title', index=False)
    
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#step 3: create output summaries for spindle density

#create path variables to use for accessing folders
archiveFolder = Path(os.getcwd()) / "Archive"
sleepScores = Path(os.getcwd()) / "sleepScorer"


#yes, the values <1,2,3> are redundant, it's just for code readibility
sleepStageMap = {"W":0, 0:0, 1:1, 2:2, 3:3, "R":4, 4:4}

for fileName in sorted(os.listdir(archiveFolder)):

    print("\n~~~~~~~~~~~~~~\n", fileName, "\n~~~~~~~~~~~~~~\n")
    try:
        logging.info("\n\n\n~~~~~~~~~~~~~~~~ " + fileName + " ~~~~~~~~~~~~~~~~")
        if fileName[-4:] != ".edf":
            logging.info("Skipping non EDF file: " + fileName)
            continue
        
        edfFileName = archiveFolder / fileName
        sleepScoreFileName = sleepScores / (fileName[:-4].replace("NS", "") + "_ODS.ods")
        sheetName = fileName[:-4]


        logging.info("Loading files: " + edfFileName.name + '\t' + sleepScoreFileName.name)
    
        raw = mne.io.read_raw_edf(edfFileName, preload = True, verbose = False)
        sf = raw.info['sfreq']
        logging.info("\tLoaded EEG data, sampling freq = " + str(sf))


        #read sleep score ods file, and convert to array of ints for yasa to process
        temp = pd.read_excel(sleepScoreFileName, engine='odf').to_numpy()
        hypno_ints = np.ndarray(temp.size)
        for i in range(temp.size):
            val = temp[i][0]
            if val in sleepStageMap:
                hypno_ints[i] = sleepStageMap[val]
            else:
                raise ValueError('Sleep Stage Mapping Not Found')


        #create first table
        hypno = yasa.hypno_upsample_to_data(hypno_ints, sf_hypno = 1/30, data = raw)

        sp_table1 = yasa.spindles_detect(data=raw, hypno=hypno, include=(0, 1, 2, 3, 4))
        summary_table1 = sp_table1.summary(grp_chan = True, grp_stage = True, aggfunc='mean').round(3)


        #create second table
        table2_start = len(summary_table1) + 5
        summary_table2 = sp_table1.summary(grp_chan=False, grp_stage=True, aggfunc='mean').round(3)

        #create third table
        table3_start = len(summary_table2) + table2_start + 5
        sp_table3 = yasa.spindles_detect(data = raw)
        summary_table3 = sp_table3.summary(grp_chan=True).round(3)
        
        with pd.ExcelWriter(outputFile, engine="openpyxl", mode='a') as writer:
            writer.if_sheet_exists = "overlay"
            if sheetName in writer.sheets.keys():
                logging.info("SHEET ALREADY EXISTS IN EXCEL FILE, OVERRIDING: " + sheetName)
            summary_table1.to_excel(writer, sheet_name = sheetName)            
            summary_table2.to_excel(writer, sheet_name = sheetName, startrow = table2_start)
            summary_table3.to_excel(writer, sheet_name = sheetName, startrow = table3_start)

    except Exception as errMessage:
        with open("errorLogFile.txt", "a") as fin:
            print(fileName + ": " + str(errMessage), file = fin)
             
    

import io
import os
import os.path
import random
from site import venv
import json
import numpy as np
import xlrd
from flask import Flask, jsonify, request, Response
from flask_cors import CORS  # for CORS error
import pyodbc  # to connect to SQL Server Database
from datetime import datetime
import dateutil.parser
import bcrypt
from bcrypt import hashpw
from datetime import datetime
import math
import pandas as pd
from numpy.lib.function_base import insert
from openpyxl import load_workbook
from pandas import DataFrame, read_csv
from database import Database
from configparser import ConfigParser
from werkzeug.datastructures import ImmutableMultiDict
from werkzeug.utils import secure_filename
# from escpos.printer import Usb # to print report

import win32print

import win32api
UPLOAD_FOLDER = f"{os.path.dirname(os.getcwd())}/Reporting_Backend/file"

app = Flask(__name__)  # creating a flask instance
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)
app.config['CORS_HEADERS'] = 'Content-Type'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


config = ConfigParser()
config.read("config.ini")
database = Database(config)

# ***************************************************************************************************************
# ******************************************ATTENDANCE MODULE IMPLEMENTATION******************************************
# ***************************************************************************************************************

@app.route('/draw', methods=['POST'])
def draw():
    conn = database.connect()
    if conn is not None:
        cursor = conn.cursor()
    if request.method == 'POST':
        reqData = request.get_json()
        print("*** Request Received - Lottery Draw***")
        draw = reqData['draw']
        ranks = ['Brig', 'Col', 'Lt Col', 'Maj', 'Capt', 'Lt']

        catFourLt = 5
        catFourCapt = 31
        catFourMaj = 40
        catFourLtCol = 0
        catFourCol = 0
        catFourBrig = 0

        catThreeLt = 8
        catThreeCapt = 24
        catThreeMaj = 32
        catThreeLtCol = 10
        catThreeCol = 0
        catThreeBrig = 0

        catTwoLt = 5
        catTwoCapt = 10
        catTwoMaj = 10
        catTwoLtCol = 11
        catTwoCol = 2
        catTwoBrig = 2

        catOneLt = 6
        catOneCapt = 10
        catOneMaj = 10
        catOneLtCol = 8
        catOneCol = 3
        catOneBrig = 3

        classOne = [catOneBrig, catOneCol, catOneLtCol, catOneMaj, catOneCapt, catOneLt]
        classTwo = [catTwoBrig, catTwoCol, catTwoLtCol, catTwoMaj, catTwoCapt, catTwoLt]
        classThree = [catThreeBrig, catThreeCol, catThreeLtCol, catThreeMaj, catThreeCapt, catThreeLt]
        classFour = [catFourBrig, catFourCol, catFourLtCol, catFourMaj,  catFourCapt, catFourLt]

        classes = [classOne, classTwo, classThree, classFour]
        catOneWinners = []
        catTwoWinners = []
        catThreeWinners = []
        catFourWinners = []
        jcoWinners = []
        sldrsWinner = []
        sldrsWinnerCatOne = []
        sldrsWinnerCatTwo = []
        sldrsWinnerCatThree = []
        sldrsWinnerCatFour = []
        sldrsWinnerCatFive = []
        param = 'paNum'

        if draw == 1:
            for classs in classes:
                cat = classes.index(classs)+1
                rnk = 0
                for v in classs:
                    rank = ranks[rnk]
                    rnk = rnk+1
                    drawQuery = f"""SELECT TOP ({v}) paNum, rank, name, unit, choiceOne, choiceTwo, grading, indexNo FROM luckydraw_tl 
                                WHERE (choiceOne = {cat} OR choiceTwo = {cat}) AND RANK = '{rank}' AND STATUS != 1 ORDER BY NEWID()"""


                    attendanceQueryExec = cursor.execute(drawQuery)
                    attendanceQueryData = cursor.fetchall()


                    for recordss in attendanceQueryData:
                        if cat == 1:
                            catOneWinners.append({"paNum": recordss[0], "rank": recordss[1], "name": recordss[2], "unit": recordss[3],
                                               "chOne": recordss[4], "chTwo": recordss[5], "grade": recordss[6], "index": recordss[7],  "cat": cat})
                        elif cat == 2:
                            catTwoWinners.append({"paNum": recordss[0], "rank": recordss[1], "name": recordss[2], "unit": recordss[3],
                                               "chOne": recordss[4], "chTwo": recordss[5], "grade": recordss[6], "index": recordss[7],  "cat": cat})
                        elif cat == 3:
                            catThreeWinners.append({"paNum": recordss[0], "rank": recordss[1], "name": recordss[2], "unit": recordss[3],
                                               "chOne": recordss[4], "chTwo": recordss[5], "grade": recordss[6], "index": recordss[7],  "cat": cat})
                        elif cat == 4:
                            catFourWinners.append({"paNum": recordss[0], "rank": recordss[1], "name": recordss[2], "unit": recordss[3],
                                               "chOne": recordss[4], "chTwo": recordss[5], "grade": recordss[6], "index": recordss[7],  "cat": cat})

                        query = f"""UPDATE luckydraw_tl SET STATUS = 1, wonInDraw = 1, wonCategory = {cat} WHERE paNum = {recordss[0]}"""
                        cursor.execute(query)

                        insertDataQuery = f"""INSERT INTO winners_tl(PaNo_ArmyNo ,Offr_JCO_Sldr , Rank, Name, Unit, Arm, Trade, wonInCat, wonInDraw)
                                                                   VALUES (?,?,?,?,?,?,?,?,?) """
                        cursor.execute(insertDataQuery, (recordss[0], 'Offr', recordss[1], recordss[2], recordss[3], 'N/A', 'N/A', cat, 1))

                    conn.commit()


            return jsonify({"response": "success", "winCatOne": catOneWinners, "winCatTwo": catTwoWinners,
                            "winCatThree": catThreeWinners, "winCatFour": catFourWinners, "cat": draw}), 200


        elif draw == 2:
            drawQueryOne = f"""SELECT TOP (10) ArmyNo, Rank, Trade, Name, Arm FROM luckyDrawJS 
                                    WHERE Trade='Clk' AND JCO_Sldr='JCO' AND STATUS != 1 ORDER BY NEWID()"""

            jcoClkQueryExec = cursor.execute(drawQueryOne)
            jcoClkQueryData = cursor.fetchall()

            for recordss in jcoClkQueryData:
                jcoWinners.append(
                    {"armyNo": recordss[0], "rank": recordss[1], "trade": recordss[2], "name": recordss[3],
                     "arm": recordss[4]})

                query = f"""UPDATE luckyDrawJS SET STATUS = 1 WHERE ArmyNo = {recordss[0]}"""
                cursor.execute(query)

                insertDataQuery = f"""INSERT INTO winners_tl(PaNo_ArmyNo ,Offr_JCO_Sldr , Rank, Name, Unit, Arm, Trade, wonInCat, wonInDraw)
                                                                                   VALUES (?,?,?,?,?,?,?,?,?) """
                cursor.execute(insertDataQuery,
                               (recordss[0], 'JCOs', recordss[1], recordss[3], 'N/A', recordss[4], recordss[2], 'N/A', 2))

            conn.commit()

            # **************************************************************************************************************

            drawQueryTwo = f"""SELECT TOP (40) ArmyNo, Rank, Trade, Name, Arm FROM luckyDrawJS 
                                                WHERE Trade='GD' AND JCO_Sldr='JCO' AND STATUS != 1 ORDER BY NEWID()"""

            jcoGdQueryExec = cursor.execute(drawQueryTwo)
            jcoGdQueryData = cursor.fetchall()

            for recordss in jcoGdQueryData:
                jcoWinners.append(
                    {"armyNo": recordss[0], "rank": recordss[1], "trade": recordss[2], "name": recordss[3],
                     "arm": recordss[4]})

                query = f"""UPDATE luckyDrawJS SET STATUS = 1 WHERE ArmyNo = {recordss[0]}"""
                cursor.execute(query)

                insertDataQuery = f"""INSERT INTO winners_tl(PaNo_ArmyNo ,Offr_JCO_Sldr , Rank, Name, Unit, Arm, Trade, wonInCat, wonInDraw)
                                                                                                   VALUES (?,?,?,?,?,?,?,?,?) """
                cursor.execute(insertDataQuery,
                               (recordss[0], 'JCO', recordss[1], recordss[3], 'N/A', recordss[4], recordss[2], 'N/A', 2))

            conn.commit()

            return jsonify({"response": "success", "winJcos": jcoWinners, "cat": 2}), 200

        elif draw == 3:
            bikeCatOne = [15, 15, 10, 10]
            bikeCatTwo = [28, 28, 16, 28]
            bikeCatThree = [27, 27, 16, 30]
            bikeCatFour = [26, 29, 16, 29]
            bikeCatFive = [26, 28, 15, 31]

            ranksSldrs = ['Hav', 'Nk', 'Lnk', 'Sep']
            allCats = [bikeCatOne, bikeCatTwo, bikeCatThree, bikeCatFour, bikeCatFive]

            catNum = 1
            for cats in allCats:
                rnk = 0
                index = 0
                for cat in cats:
                    print("Cat Number",catNum)
                    print("Cat Count",cat)
                    rank = ranksSldrs[rnk]
                    rnk = rnk + 1
                    print("Rank", rank)

                    print(catNum)
                    drawQueryOne = f"""SELECT TOP ({cat}) ArmyNo, Rank, Trade, Name, Arm FROM luckyDrawJS 
                                            WHERE Rank='{rank}' AND JCO_Sldr='Sldr' AND STATUS != 1 ORDER BY NEWID()"""

                    print(drawQueryOne)

                    sldrQueryExec = cursor.execute(drawQueryOne)
                    sldrQueryData = cursor.fetchall()


                    for records in sldrQueryData:
                        index = index + 1
                        if catNum == 1:
                            sldrsWinnerCatOne.append(
                                {"armyNo": records[0], "rank": records[1], "trade": records[2], "name": records[3],
                                 "arm": records[4], "categ": catNum, "index": index})
                        elif catNum == 2:
                            sldrsWinnerCatTwo.append(
                                {"armyNo": records[0], "rank": records[1], "trade": records[2], "name": records[3],
                                 "arm": records[4], "categ": catNum, "index": index})
                        elif catNum == 3:
                            sldrsWinnerCatThree.append(
                                {"armyNo": records[0], "rank": records[1], "trade": records[2], "name": records[3],
                                 "arm": records[4], "categ": catNum, "index": index})
                        elif catNum == 4:
                            sldrsWinnerCatFour.append(
                                {"armyNo": records[0], "rank": records[1], "trade": records[2], "name": records[3],
                                 "arm": records[4], "categ": catNum, "index": index})
                        elif catNum == 5:
                            sldrsWinnerCatFive.append(
                                {"armyNo": records[0], "rank": records[1], "trade": records[2], "name": records[3],
                                 "arm": records[4], "categ": catNum, "index": index})

                        query = f"""UPDATE luckyDrawJS SET STATUS = 1 WHERE ArmyNo = {records[0]}"""
                        cursor.execute(query)

                        insertDataQuery = f"""INSERT INTO winners_tl(PaNo_ArmyNo ,Offr_JCO_Sldr , Rank, Name, Unit, Arm, Trade, wonInCat, wonInDraw)
                                                                                                           VALUES (?,?,?,?,?,?,?,?,?) """
                        cursor.execute(insertDataQuery,
                                       (records[0], 'Sldr', records[1], records[3], 'N/A', records[4], records[2], catNum ,3))

                    conn.commit()
                catNum = catNum + 1
            sldrsWinner.append(sldrsWinnerCatOne)
            sldrsWinner.append(sldrsWinnerCatTwo)
            sldrsWinner.append(sldrsWinnerCatThree)
            sldrsWinner.append(sldrsWinnerCatFour)
            sldrsWinner.append(sldrsWinnerCatFive)

            return jsonify({"response": "success", "winSldrs": sldrsWinner, "cat": 3}), 200




@app.route('/reset', methods=['GET'])
def reset():
    conn = database.connect()
    if conn is not None:
        cursor = conn.cursor()
    if request.method == 'GET':
        query = f"""UPDATE luckydraw_tl SET STATUS = 0,  wonInDraw = NULL, wonCategory = NULL"""
        cursor.execute(query)
        conn.commit()

        query = f"""UPDATE luckyDrawJS SET STATUS = 0"""
        cursor.execute(query)
        conn.commit()

        query = f"""TRUNCATE TABLE winners_tl"""
        cursor.execute(query)
        conn.commit()
        return jsonify({"response": "success"}), 200


@app.route('/printPDF', methods=['POST'])
def printPDF():
    if request.method == 'POST':
        print(request.form)
        file = request.files['pdf']
        if file:
            file.save(os.path.join('./', f'{request.form["filename"]}.pdf'))

            GHOSTSCRIPT_PATH = "./GHOSTSCRIPT\\bin\\gswin32.exe"
            GSPRINT_PATH = "./GSPRINT\\gsprint.exe"

            # YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT
            currentprinter = win32print.GetDefaultPrinter()

            win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH+'" -color -printer "'+currentprinter+'" "'+ f'{request.form["filename"]}.pdf' + '', '.', 0)
        return jsonify({"response": "success"}), 200




if __name__ == '__main__':
    app.run(host='127.0.0.1', port=8090, debug=False)
    # app.run(port=8090, debug=True)

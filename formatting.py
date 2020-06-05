from time import sleep
def format_sheet(sheetName, sheetId, date, client):
    ss = client.open(sheetName)

    #starting column number for merging first row cells
    merge_head_start_col = 0
    #starting column number for color gradient
    color_start_col = 1
    #starting column number for merging ratio row cells
    merge_end_start_col = 0
    for i in range(0, 32):
        sleep(1) #to avoid error of exceeding 100 requests within 100s
        merge_head_end_col = merge_head_start_col+4
        color_end_col = color_start_col+3
        merge_end_end_col = merge_end_start_col+2
        # print(startcol, endcol)
        body = {
            "requests": [
                {
                    "mergeCells": { #merge first row
                        "mergeType": "MERGE_ALL",
                        "range": {
                            "sheetId": sheetId,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": merge_head_start_col,
                            "endColumnIndex": merge_head_end_col
                        }
                    }
                }, {
                    'addConditionalFormatRule': { #color gradient
                        'rule': {
                            'ranges': [
                                {
                                    'sheetId': sheetId,
                                    'startRowIndex': 2,
                                    'endRowIndex': 91,
                                    'startColumnIndex': color_start_col,
                                    'endColumnIndex': color_end_col,
                                }
                            ],
                            'gradientRule': { #currently gradient from white to pinkish red
                                "minpoint": {
                                    "color": { #color of min point
                                        "red": 1,
                                        "green": 1,
                                        "blue": 1,
                                    },
                                    "type": "MIN"
                                },
                                "maxpoint": {
                                    "color": { #color of max point
                                        "red": 0.901,
                                        "green": 0.486,
                                        "blue": 0.450,
                                    },
                                    "type": "MAX"
                                }
                            }
                        },
                        'index': 0
                    }
                }, {
                    "mergeCells": { #merge ratio row for "Put/Call Ratio"
                        "mergeType": "MERGE_ALL",
                        "range": {
                            "sheetId": sheetId,
                            "startRowIndex": 95,
                            "endRowIndex": 96,
                            "startColumnIndex": merge_end_start_col,
                            "endColumnIndex": merge_end_end_col
                        }
                    }
                }, {
                    "mergeCells": { #merge ratio row for number
                        "mergeType": "MERGE_ALL",
                        "range": {
                            "sheetId": sheetId,
                            "startRowIndex": 95,
                            "endRowIndex": 96,
                            "startColumnIndex": merge_end_start_col+2,
                            "endColumnIndex": merge_end_end_col+2
                        }
                    }
                }, {
                    "updateSheetProperties": { #change name of sheet
                        "properties": {"title": date},
                        "fields": "title"
                    }
                }
            ]
        }
        res = ss.batch_update(body)
        merge_head_start_col+=4
        color_start_col+=4
        merge_end_start_col+=4


    sheet = ss.sheet1
    sheet.format("A1:DX1", { #color and border for the first row
        "backgroundColor": { #yellow
        "red": 0.988,
        "green": 0.956,
        "blue": 0.639
        },
        "borders": {
            "top": {
                "style": "SOLID_MEDIUM",
            },
            "bottom": {
                "style": "SOLID_MEDIUM",
            },
            "left": {
                "style": "SOLID_MEDIUM",
            },
            "right": {
                "style": "SOLID_MEDIUM",
            }
        },
    })
    sheet.format("A1:DZ100", { #center align all text
        "horizontalAlignment": "CENTER",
    })

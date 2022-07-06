{
    "actions": [
        {
            "id": "executeWriteData",
            "type": "ExecuteFunction",
            "functionName": "writeData"
        }
    ],
    "tabs": [
        {
            "id": "CtxTab1",
            "label": "Contoso Data",
            "groups": [
                {
                    "id": "CustomGroup111",
                    "label": "Insertion",
                    "icon": [
                        {
                            "size": 16,
                            "sourceLocation":"../../assets/icon-16.png"
                        },
                        {
                            "size": 32,
                            "sourceLocation":"../../assets/icon-32.png"
                        },
                        {
                            "size": 80,
                            "sourceLocation":"../../assets/icon-80.png"
                        }
                    ],
                    "controls": [
                        {
                            "type": "Button",
                            "id": "CtxBt112",
                            "actionId": "executeWriteData",
                            "enabled": false,
                            "label": "Write Data",
                            "superTip": {
                                "title": "Data Insertion",
                                "description": "Use this button to insert data into the document."
                            },
                            "icon": [
                                {
                                    "size": 16,
                                    "sourceLocation":"../../assets/icon-16.png"
                                },
                                {
                                    "size": 32,
                                    "sourceLocation":"../../assets/icon-32.png"
                                },
                                {
                                    "size": 80,
                                    "sourceLocation":"../../assets/icon-80.png"
                                }
                            ],
                        }
                    ]
                }
            ]
          }
    ],
}
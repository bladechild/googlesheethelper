const expect = require('expect');
const {CreateSpreadSheet} = require('../handler');
describe('Create Sheet',()=>{
    it('should create a google sheet',(done)=>{
        const sheets = [
            {
                title: 'Sheet Test 1',
                rowCount: 10,
                columnCount: 10,
                frozenRowCount: 1,
                frozenColumnCount: 2,
                hideGridlines: false,
                data:[
                    {
                        startRow:2,
                        startColumn:3,
                        rowData:[
                            {
                                value:2
                            },
                            {
                                value:6,
                                note: 'Test note 1'
                            },
                            {
                                value: 'Test value string'
                            },
                            {
                                value: true
                            }
                        ]
                    },
                    {
                        startRow:5,
                        startColumn:1,
                        rowData:[
                            {
                                value:3
                            },
                            {
                                value:5,
                                note: 'Test note 2'
                            },
                            {
                                value: 'Test value2 string'
                            },
                            {
                                value: false
                            }
                        ]
                    }
                ]
            },
            {
                title: 'Sheet Test 2',
                rowCount: 5,
                columnCount: 5,
                frozenRowCount: 4,
                frozenColumnCount: 4,
                hideGridlines: true,
                data:[
                    {
                        startRow:2,
                        startColumn:0,
                        rowData:[
                            {
                                value:2
                            },
                            {
                                value:6,
                                note: 'Test note 1'
                            },
                            {
                                value: 'Test value string'
                            },
                            {
                                value: true
                            }
                        ]
                    },
                    {
                        startRow:4,
                        startColumn:1,
                        rowData:[
                            {
                                value:3
                            },
                            {
                                value:5,
                                note: 'Test note 2'
                            },
                            {
                                value: 'Test value2 string'
                            },
                            {
                                value: false
                            }
                        ]
                    }
                ]
            }
        ];
        CreateSpreadSheet('Test Sheet',sheets,(error,response)=>{
            if(error)
                done(error);
            else
                done();
        });
    });
})
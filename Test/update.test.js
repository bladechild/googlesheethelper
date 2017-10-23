const expect = require('expect');
const {UpdateSpreadSheet,UpdateSheetProperties,UpdateRows} = require('../handler');
describe('Update Spread Sheet',()=>{
    it('should update a google sheet',(done)=>{
        UpdateSpreadSheet('1X9kRo2dAxW0uXurfPksnv1upcaa2A09ZSXKOd5WXPJU','My modified title2',(error,response)=>{
            if(error)
                done(error);
            else
            {
                console.log(response);
                expect(response.spreadsheetId).toBe('1X9kRo2dAxW0uXurfPksnv1upcaa2A09ZSXKOd5WXPJU');
                done();
            }
        });
    });

    it('should update a sheet',(done)=>{
        UpdateSheetProperties('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI','My Sheet 4','My Sheet 4',9,9,1,1,false,(error,response)=>{
            if(error)
                done(error);
            else
            {
                console.log(response);
                expect(response.spreadsheetId).toBe('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI');
                done();
            }
        });
    });
    it('should update rows',(done)=>{
        UpdateRows('1JWEARq8pIbwoU-AZfvFq58w4oflO6tOpS4-8AClPyug','Sheet Test 1',[[1,2,3,'=SUM(D6:F6)'],[4,5,6,45],[7,8,9,10]],1,3,(error,response)=>{
            if(error)
                done(error);
            else
            {
                console.log(response);
                expect(response.spreadsheetId).toBe('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI');
                done();
            }
        });
    });
})
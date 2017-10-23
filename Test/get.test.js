const expect = require('expect');
const {GetSpreadSheet} = require('../handler');
describe('Get Sheet',()=>{
    it('should get a google sheet',(done)=>{
        GetSpreadSheet('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI',(error,response)=>{
            if(error)
                done(error);
            else
            {
                console.log(JSON.stringify(response));
                expect(response.spreadsheetId).toBe('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI');
                done();
            }
        });
    });
})
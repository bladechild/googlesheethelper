const expect = require('expect');
const {AddSheet} = require('../handler');
describe('Add a Sheet',()=>{
    it('should add a sheet',(done)=>{
        AddSheet('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI','My Sheet 5',2,2,1,1,false,(error,response)=>{
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
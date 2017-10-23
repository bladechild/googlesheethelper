const expect = require('expect');
const {DeleteSheet} = require('../handler');
describe('Delete a Sheet',()=>{
    it('should delete a sheet',(done)=>{
        DeleteSheet('13WZWyDuSZvd_MhR8UfruD8xy0xIaJ1F45hW16Pki9ZI','My Sheet 5',(error,response)=>{
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
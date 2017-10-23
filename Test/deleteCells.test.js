const expect = require('expect');
const {DeleteRangedCells} = require('../handler');
describe('Delete cells',()=>{
    it('should delete cells',(done)=>{
        DeleteRangedCells('1JWEARq8pIbwoU-AZfvFq58w4oflO6tOpS4-8AClPyug','Sheet Test 1',null,null,null,null,'COLUMNS',(error,response)=>{
            if(error)
                done(error);
            else
            {
                console.log(response);
                expect(response.spreadsheetId).toBe('1JWEARq8pIbwoU-AZfvFq58w4oflO6tOpS4-8AClPyug');
                done();
            }
        });
    });
})
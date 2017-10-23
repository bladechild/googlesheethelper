const expect = require('expect');
const {AddRows} = require('../handler');
describe('Add a Row',()=>{
    it('should add a row',(done)=>{
        AddRows('1JWEARq8pIbwoU-AZfvFq58w4oflO6tOpS4-8AClPyug','Sheet Test 1',[[2,2,1,1]],(error,response)=>{
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
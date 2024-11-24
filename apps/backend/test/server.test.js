const request = require('supertest');
const app = require('../server.js'); // CommonJS import

// Dynamically import chai
let chai, expect;

before(async () => {
    chai = await import('chai'); // Use dynamic import
    expect = chai.expect;
});

describe('Server Tests', () => {
    it('should return status 200 on GET /', (done) => {
        request(app)
            .get('/')
            .expect(200)
            .end(done);
    });

    it('should return data for /undyed-yarn', (done) => {
        request(app)
            .get('/undyed-yarn')
            .expect(200)
            .expect((res) => {
                expect(res.body).to.be.an('array');
            })
            .end(done);
    });
});

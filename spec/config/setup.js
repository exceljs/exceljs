const chai = require('chai');

const chaiXml = require('chai-xml');
const chaiDatetime = require('chai-datetime');
const dirtyChai = require('dirty-chai');

chai.use(chaiXml);
chai.use(chaiDatetime);
chai.use(dirtyChai);

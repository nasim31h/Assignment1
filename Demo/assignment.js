const openpyxl = require('openpyxl');
const webdriver = require('selenium-webdriver');
const { By, Key } = require('selenium-webdriver');
const { Options } = require('selenium-webdriver/chrome');
const { Builder, Capabilities } = require('selenium-webdriver');
const { parse } = require('node-html-parser');
const fs = require('fs');
const path = require('path');


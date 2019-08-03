

// Save to SharePoint Assets Library
var spsave = require("spsave").spsave;
var coreOptions = {
    siteUrl: 'http://fkrea/sites/documentation',
    notification: true,
    checkin: true,
    checkinType: 1,
    glob: ['dist/js/*.*', 'dist/css/*.css']
};
var creds = {
    username: 'spapp',
    password: 'Fkr2016',
    domain: 'fond-mos'
};

var fileOptions = {
    folder: 'SitePages/js/upload',
    glob: ['dist/js/*.*']
    // fileName: 'dist/js/app.js'
};

var fileOptionsCss = {
    folder: 'Style Library/fkrea/upload',
    glob: ['dist/css/*.*']
    // fileName: 'dist/js/app.js'
};

// vue.config.js
module.exports = {
    filenameHashing: false,
    configureWebpack: {
        output: {
            libraryTarget: 'var',
            globalObject: 'this',
            library: 'dynamicapp'
        },
        plugins: [
            {
                apply: (compiler) => {
                    compiler.hooks.afterEmit.tap('AfterEmitPlugin', (compilation) => {
                        spsave(coreOptions, creds, fileOptions).then(() => { }).catch((e) => {
                            console.error(e);
                        });
                        spsave(coreOptions, creds, fileOptionsCss).then(() => { }).catch((e) => {
                            console.error(e);
                        });
                    });
                }
            }
        ]
    }
}
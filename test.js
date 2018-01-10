const cp = require('child_process');
test();

function test() {
    cp.exec('node action.js',function (err, stdout, stderr) {
        if(err){
            console.log('stderr: ' + stderr);
        } else {
            console.log('stdout: ' + stdout);
        }
    });
}
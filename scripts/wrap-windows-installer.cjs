const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const makeDir = path.resolve(__dirname, '../artifacts/forge-out/make/squirrel.windows/x64');
const setup = path.join(makeDir, 'StockReporterAssistantSetup.exe');
const internal = path.join(makeDir, 'StockReporterAssistantSquirrelSetup.exe');
const sourceSetup = process.argv.includes('--input-existing-squirrel') ? internal : setup;
if (process.platform !== 'win32') throw new Error('Installer wrapper requires Windows.');
if (!fs.existsSync(sourceSetup)) throw new Error('Squirrel Setup.exe is missing; run Electron Forge make first.');

const stagingParent = fs.realpathSync(os.tmpdir());
const staging = fs.mkdtempSync(path.join(stagingParent, 'reporter-installer-wrap-'));
try {
  const wrapped = path.join(staging, 'StockReporterAssistantSetup.exe');
  const embeddedSetup = path.join(staging, 'StockReporterAssistantSquirrelSetup.exe');
  fs.copyFileSync(sourceSetup, embeddedSetup);
  for (const name of ['installer-preflight.ps1', 'installer-preflight-core.ps1']) {
    const source = fs.readFileSync(path.join(__dirname, name), 'utf8');
    fs.writeFileSync(path.join(staging, name), `\ufeff${source}`, 'utf8');
  }
  const sed = [
    '[Version]',
    'Class=IEXPRESS',
    'SEDVersion=3',
    '[Options]',
    'PackagePurpose=InstallApp',
    'ShowInstallProgramWindow=0',
    'HideExtractAnimation=1',
    'UseLongFileName=1',
    'InsideCompressed=1',
    'CAB_FixedSize=0',
    'CAB_ResvCodeSigning=0',
    'RebootMode=N',
    'InstallPrompt=%InstallPrompt%',
    'DisplayLicense=%DisplayLicense%',
    'FinishMessage=%FinishMessage%',
    'TargetName=%TargetName%',
    'FriendlyName=%FriendlyName%',
    'AppLaunched=%AppLaunched%',
    'PostInstallCmd=%PostInstallCmd%',
    'AdminQuietInstCmd=%AdminQuietInstCmd%',
    'UserQuietInstCmd=%UserQuietInstCmd%',
    'SourceFiles=SourceFiles',
    '[Strings]',
    'InstallPrompt=',
    'DisplayLicense=',
    'FinishMessage=',
    `TargetName=${wrapped}`,
    'FriendlyName=Stock Reporter Assistant Setup',
    'AppLaunched=powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File installer-preflight.ps1',
    'PostInstallCmd=<None>',
    'AdminQuietInstCmd=',
    'UserQuietInstCmd=',
    'FILE0="installer-preflight.ps1"',
    'FILE1="installer-preflight-core.ps1"',
    'FILE2="StockReporterAssistantSquirrelSetup.exe"',
    '[SourceFiles]',
    `SourceFiles0=${staging}\\`,
    '[SourceFiles0]',
    '%FILE0%=',
    '%FILE1%=',
    '%FILE2%=',
  ].join('\r\n') + '\r\n';
  const sedPath = path.join(staging, 'package.sed');
  fs.writeFileSync(sedPath, sed, 'ascii');
  console.log('Building prompted installer from Squirrel package...');
  execFileSync(path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'iexpress.exe'), ['/N', '/Q', sedPath], { stdio: 'inherit', timeout: 600000 });
  if (!fs.existsSync(wrapped)) throw new Error('IExpress did not create the prompted installer.');
  fs.copyFileSync(embeddedSetup, internal);
  fs.copyFileSync(wrapped, setup);
  console.log('Interactive installer preflight wrapped Squirrel Setup.exe.');
} finally {
  const resolvedStaging = fs.realpathSync(staging);
  if (!resolvedStaging.startsWith(`${stagingParent}${path.sep}`)) {
    throw new Error('Installer staging escaped the expected temporary directory.');
  }
  try {
    fs.rmSync(resolvedStaging, { recursive: true, force: true, maxRetries: 5, retryDelay: 250 });
  } catch (error) {
    console.warn(`Installer staging could not be removed yet: ${error.code || error.message}`);
  }
}

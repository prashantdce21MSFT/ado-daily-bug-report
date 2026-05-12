const vscode = require('vscode');
const fs = require('fs');
const path = require('path');

function readTpl(ctx, rel) {
  return fs.readFileSync(path.join(ctx.extensionPath, rel), 'utf8');
}

function ensureDir(p) { fs.mkdirSync(p, { recursive: true }); }

async function ask(prompt, value, placeHolder) {
  return vscode.window.showInputBox({ prompt, value, placeHolder, ignoreFocusOut: true });
}

async function scaffold(ctx) {
  const folderPick = await vscode.window.showOpenDialog({
    canSelectFolders: true, canSelectFiles: false, canSelectMany: false,
    openLabel: 'Select parent folder for new report'
  });
  if (!folderPick || !folderPick.length) return;
  const parent = folderPick[0].fsPath;

  const name = await ask('Folder name for the new report', 'MyTeam ADO Daily Report');
  if (!name) return;
  const target = path.join(parent, name);
  if (fs.existsSync(target)) {
    const ow = await vscode.window.showWarningMessage(`${target} exists. Overwrite files?`, 'Yes', 'No');
    if (ow !== 'Yes') return;
  }
  ensureDir(target);
  ensureDir(path.join(target, 'logs'));

  const org = await ask('ADO Organization URL', 'https://dev.azure.com/your-org');
  if (!org) return;
  const proj = await ask('ADO Project name', 'YourProject');
  if (!proj) return;
  const plan = await ask('Test Plan ID (number)', '12345');
  if (!plan) return;
  const iter = await ask('Iteration path (backslash separated)', `${proj}\\Iteration 1`);
  if (!iter) return;
  const iterStart = await ask('Iteration start date YYYY-MM-DD (optional, blank to skip)', '');
  const recipients = await ask('Recipients (Outlook format: "Name <email>; Name <email>")', 'You <you@example.com>');
  const alert = await ask('Alert recipient (for auth/failure pings)', 'you@example.com');
  const subjectTag = await ask('Subject line tag', '[ADO Daily]');

  const subs = {
    '__OUT__': target.replace(/\\/g, '\\\\'),
    '__ORG__': org,
    '__PROJ__': proj,
    '__PLAN__': plan,
    '__ITER__': iter.replace(/\\/g, '\\\\'),
    '__ITER_START__': iterStart || '',
    '__RECIPIENTS__': recipients,
    '__ALERT__': alert,
    '__SUBJECT_TAG__': subjectTag,
    '__WORK__': target
  };

  const apply = (s) => Object.keys(subs).reduce((acc, k) => acc.split(k).join(subs[k]), s);

  const files = [
    ['templates/daily_report.py', 'daily_report.py'],
    ['templates/run_daily.ps1', 'run_daily.ps1'],
    ['templates/open_defects_report.py', 'open_defects_report.py'],
    ['templates/run_open_defects.ps1', 'run_open_defects.ps1'],
    ['templates/SETUP.md', 'SETUP.md']
  ];
  for (const [src, dst] of files) {
    const content = apply(readTpl(ctx, src));
    fs.writeFileSync(path.join(target, dst), content, 'utf8');
  }

  const open = await vscode.window.showInformationMessage(
    `Scaffolded at ${target}`, 'Open Folder', 'Open SETUP.md'
  );
  if (open === 'Open Folder') {
    vscode.commands.executeCommand('vscode.openFolder', vscode.Uri.file(target), { forceNewWindow: true });
  } else if (open === 'Open SETUP.md') {
    const doc = await vscode.workspace.openTextDocument(path.join(target, 'SETUP.md'));
    vscode.window.showTextDocument(doc);
  }
}

async function installPrompts(ctx) {
  const ws = vscode.workspace.workspaceFolders;
  if (!ws || !ws.length) {
    vscode.window.showErrorMessage('Open a workspace folder first.');
    return;
  }
  const target = path.join(ws[0].uri.fsPath, '.github', 'prompts');
  ensureDir(target);
  const promptsDir = path.join(ctx.extensionPath, 'prompts');
  const files = fs.readdirSync(promptsDir);
  let n = 0;
  for (const f of files) {
    fs.copyFileSync(path.join(promptsDir, f), path.join(target, f));
    n++;
  }
  vscode.window.showInformationMessage(`Installed ${n} prompt files to ${target}`);
}

async function registerTask() {
  const work = await ask('Path to your scaffolded report folder', 'C:\\MyTeam ADO Daily Report');
  if (!work) return;
  const taskName = await ask('Scheduled task name', 'My ADO Daily Bug Report');
  if (!taskName) return;
  const time = await ask('Time of day (e.g. 7:00AM)', '7:00AM');
  if (!time) return;

  const cmd = [
    `$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "${work}\\run_daily.ps1"'`,
    `$trigger = New-ScheduledTaskTrigger -Daily -At '${time}'`,
    `$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)`,
    `Register-ScheduledTask -TaskName '${taskName}' -Action $action -Trigger $trigger -Settings $settings`
  ].join('\n');

  await vscode.env.clipboard.writeText(cmd);
  const doc = await vscode.workspace.openTextDocument({ language: 'powershell', content: cmd });
  vscode.window.showTextDocument(doc);
  vscode.window.showInformationMessage('Register-ScheduledTask command copied to clipboard.');
}

async function runNow() {
  const work = await ask('Path to your scaffolded report folder', 'C:\\MyTeam ADO Daily Report');
  if (!work) return;
  const term = vscode.window.createTerminal({ name: 'ADO Daily Report (manual run)', cwd: work });
  term.show();
  term.sendText(`$env:ADO_TOKEN = & az account get-access-token --resource '499b84ac-1321-427f-aa17-267ca6975798' --query accessToken -o tsv`);
  term.sendText(`python "${work}\\daily_report.py"`);
}

function activate(context) {
  context.subscriptions.push(
    vscode.commands.registerCommand('adoDailyReport.scaffold', () => scaffold(context)),
    vscode.commands.registerCommand('adoDailyReport.installPrompts', () => installPrompts(context)),
    vscode.commands.registerCommand('adoDailyReport.registerTask', registerTask),
    vscode.commands.registerCommand('adoDailyReport.runNow', runNow)
  );
}

function deactivate() {}

module.exports = { activate, deactivate };

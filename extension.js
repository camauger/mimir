const vscode = require('vscode');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
  // Commande pour convertir un DOCX en HTML
  let convertToHtml = vscode.commands.registerCommand('docx-json-converter.convertToHtml', async (fileUri) => {
    if (!fileUri) {
      // Pas de fichier sélectionné, demander à l'utilisateur
      const docxFiles = await vscode.workspace.findFiles('**/*.docx');
      if (docxFiles.length === 0) {
        vscode.window.showErrorMessage('Aucun fichier DOCX trouvé dans l\'espace de travail.');
        return;
      }

      const items = docxFiles.map(file => ({
        label: path.basename(file.fsPath),
        description: file.fsPath,
        file: file
      }));

      const selected = await vscode.window.showQuickPick(items, {
        placeHolder: 'Sélectionnez un fichier DOCX à convertir'
      });

      if (!selected) return;
      fileUri = selected.file;
    }

    // Obtenir les chemins
    const docxPath = fileUri.fsPath;
    const outputDir = path.dirname(docxPath);
    const baseName = path.basename(docxPath, '.docx');
    const outputHtml = path.join(outputDir, `${baseName}.html`);

    // Obtenir la configuration
    const config = vscode.workspace.getConfiguration('docxJsonConverter');
    const pythonPath = config.get('pythonPath');
    const scriptPath = config.get('scriptPath');
    const saveImagesToDisk = config.get('saveImagesToDisk');

    // Construire la commande
    const saveImagesFlag = saveImagesToDisk ? '--save-images' : '--base64-images';
    const command = `"${pythonPath}" "${scriptPath}" "${docxPath}" --output-dir "${outputDir}" ${saveImagesFlag} --html`;

    // Afficher un message de progression
    vscode.window.withProgress({
      location: vscode.ProgressLocation.Notification,
      title: `Conversion de ${baseName}.docx en HTML`,
      cancellable: false
    }, async (progress) => {
      return new Promise((resolve, reject) => {
        exec(command, (error, stdout, stderr) => {
          if (error) {
            vscode.window.showErrorMessage(`Erreur lors de la conversion: ${error.message}`);
            console.error(`Erreur: ${error.message}`);
            console.error(`Sortie stderr: ${stderr}`);
            reject(error);
            return;
          }

          // Afficher un message de succès avec un bouton pour ouvrir le fichier
          vscode.window.showInformationMessage(
            `Conversion réussie: ${baseName}.html`,
            'Ouvrir'
          ).then(selection => {
            if (selection === 'Ouvrir') {
              const htmlUri = vscode.Uri.file(outputHtml);
              vscode.commands.executeCommand('vscode.open', htmlUri);
            }
          });

          resolve();
        });
      });
    });
  });

  // Commande pour convertir un DOCX en JSON
  let convertToJson = vscode.commands.registerCommand('docx-json-converter.convertToJson', async (fileUri) => {
    if (!fileUri) {
      // Pas de fichier sélectionné, demander à l'utilisateur
      const docxFiles = await vscode.workspace.findFiles('**/*.docx');
      if (docxFiles.length === 0) {
        vscode.window.showErrorMessage('Aucun fichier DOCX trouvé dans l\'espace de travail.');
        return;
      }

      const items = docxFiles.map(file => ({
        label: path.basename(file.fsPath),
        description: file.fsPath,
        file: file
      }));

      const selected = await vscode.window.showQuickPick(items, {
        placeHolder: 'Sélectionnez un fichier DOCX à convertir'
      });

      if (!selected) return;
      fileUri = selected.file;
    }

    // Obtenir les chemins
    const docxPath = fileUri.fsPath;
    const outputDir = path.dirname(docxPath);
    const baseName = path.basename(docxPath, '.docx');
    const outputJson = path.join(outputDir, `${baseName}.json`);

    // Obtenir la configuration
    const config = vscode.workspace.getConfiguration('docxJsonConverter');
    const pythonPath = config.get('pythonPath');
    const scriptPath = config.get('scriptPath');
    const saveImagesToDisk = config.get('saveImagesToDisk');

    // Construire la commande
    const saveImagesFlag = saveImagesToDisk ? '--save-images' : '--base64-images';
    const command = `"${pythonPath}" "${scriptPath}" "${docxPath}" --output-dir "${outputDir}" ${saveImagesFlag} --json`;

    // Afficher un message de progression
    vscode.window.withProgress({
      location: vscode.ProgressLocation.Notification,
      title: `Conversion de ${baseName}.docx en JSON`,
      cancellable: false
    }, async (progress) => {
      return new Promise((resolve, reject) => {
        exec(command, (error, stdout, stderr) => {
          if (error) {
            vscode.window.showErrorMessage(`Erreur lors de la conversion: ${error.message}`);
            console.error(`Erreur: ${error.message}`);
            console.error(`Sortie stderr: ${stderr}`);
            reject(error);
            return;
          }

          // Afficher un message de succès avec un bouton pour ouvrir le fichier
          vscode.window.showInformationMessage(
            `Conversion réussie: ${baseName}.json`,
            'Ouvrir'
          ).then(selection => {
            if (selection === 'Ouvrir') {
              const jsonUri = vscode.Uri.file(outputJson);
              vscode.commands.executeCommand('vscode.open', jsonUri);
            }
          });

          resolve();
        });
      });
    });
  });

  // Commande pour prévisualiser un DOCX en HTML
  let previewHtml = vscode.commands.registerCommand('docx-json-converter.previewHtml', async (fileUri) => {
    if (!fileUri) {
      // Pas de fichier sélectionné, demander à l'utilisateur
      const docxFiles = await vscode.workspace.findFiles('**/*.docx');
      if (docxFiles.length === 0) {
        vscode.window.showErrorMessage('Aucun fichier DOCX trouvé dans l\'espace de travail.');
        return;
      }

      const items = docxFiles.map(file => ({
        label: path.basename(file.fsPath),
        description: file.fsPath,
        file: file
      }));

      const selected = await vscode.window.showQuickPick(items, {
        placeHolder: 'Sélectionnez un fichier DOCX à prévisualiser'
      });

      if (!selected) return;
      fileUri = selected.file;
    }

    // Obtenir les chemins
    const docxPath = fileUri.fsPath;
    const outputDir = path.dirname(docxPath);
    const baseName = path.basename(docxPath, '.docx');

    // Créer un répertoire temporaire pour la prévisualisation
    const tempDir = path.join(outputDir, '.preview');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir);
    }
    const tempHtml = path.join(tempDir, `${baseName}.html`);

    // Obtenir la configuration
    const config = vscode.workspace.getConfiguration('docxJsonConverter');
    const pythonPath = config.get('pythonPath');
    const scriptPath = config.get('scriptPath');

    // Toujours utiliser base64 pour la prévisualisation pour éviter les problèmes de chemin
    const command = `"${pythonPath}" "${scriptPath}" "${docxPath}" --output-dir "${tempDir}" --base64-images --html`;

    // Afficher un message de progression
    vscode.window.withProgress({
      location: vscode.ProgressLocation.Notification,
      title: `Prévisualisation de ${baseName}.docx`,
      cancellable: false
    }, async (progress) => {
      return new Promise((resolve, reject) => {
        exec(command, (error, stdout, stderr) => {
          if (error) {
            vscode.window.showErrorMessage(`Erreur lors de la prévisualisation: ${error.message}`);
            console.error(`Erreur: ${error.message}`);
            console.error(`Sortie stderr: ${stderr}`);
            reject(error);
            return;
          }

          // Lire le fichier HTML généré
          fs.readFile(tempHtml, 'utf8', (err, data) => {
            if (err) {
              vscode.window.showErrorMessage(`Erreur lors de la lecture du HTML: ${err.message}`);
              reject(err);
              return;
            }

            // Créer une webview pour la prévisualisation
            const panel = vscode.window.createWebviewPanel(
              'docxPreview',
              `Prévisualisation - ${baseName}.docx`,
              vscode.ViewColumn.Two,
              {
                enableScripts: true,
                localResourceRoots: [vscode.Uri.file(tempDir)]
              }
            );

            // Définir le contenu HTML
            panel.webview.html = data;

            resolve();
          });
        });
      });
    });
  });

  context.subscriptions.push(convertToHtml);
  context.subscriptions.push(convertToJson);
  context.subscriptions.push(previewHtml);
}

function deactivate() {}

module.exports = {
  activate,
  deactivate
};
/**
 * Outils d'aide pour l'interaction avec les API d'Office.js (Word et Excel).
 */
class OfficeHelpers {
  /**
   * Détecte l'application hôte actuelle.
   * @returns {'Word' | 'Excel' | 'Browser'}
   */
  getHost() {
    if (typeof Office !== "undefined" && Office.context) {
      if (Office.context.host === Office.HostType.Word) return "Word";
      if (Office.context.host === Office.HostType.Excel) return "Excel";
    }
    return "Browser";
  }

  /**
   * Vérifie si Office.js est initialisé.
   * @returns {boolean}
   */
  isAvailable() {
    return typeof Office !== "undefined" && Office.context !== undefined;
  }

  // =========================================================================
  // ACTIONS WORD
  // =========================================================================

  /**
   * Récupère le texte actuellement sélectionné dans Word.
   * @returns {Promise<string>} Le texte sélectionné ou une chaîne vide.
   */
  async getSelectedTextWord() {
    if (this.getHost() !== "Word") return "";

    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        resolve(selection.text || "");
      }).catch((error) => {
        console.error("[OfficeHelpers] Erreur getSelectedTextWord:", error);
        reject(error);
      });
    });
  }

  /**
   * Remplace ou insère du texte par rapport à la sélection Word.
   * @param {string} text - Le texte à insérer.
   * @param {'replace' | 'after' | 'before'} location - Emplacement de l'insertion.
   * @param {boolean} trackChanges - Si vrai, applique les modifications en mode révision (suivi des modifications).
   */
  async insertTextWord(text, location = "replace", trackChanges = false) {
    if (this.getHost() !== "Word") return;

    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      let wordLocation;
      
      switch (location) {
        case "after":
          wordLocation = Word.InsertLocation.after;
          break;
        case "before":
          wordLocation = Word.InsertLocation.before;
          break;
        case "replace":
        default:
          wordLocation = Word.InsertLocation.replace;
          break;
      }

      // Sauvegarder la mise en forme de police avant de modifier le texte
      let fontBackup = {};
      try {
        selection.load("font");
        await context.sync();
        
        if (selection.font) {
          const fontProperties = ["name", "size", "color", "bold", "italic", "underline", "strikeThrough"];
          fontProperties.forEach(prop => {
            const val = selection.font[prop];
            // Si la valeur est bien définie et n'est pas un mélange ("mixed" / null)
            if (val !== null && val !== undefined && val !== "mixed" && val !== "") {
              fontBackup[prop] = val;
            }
          });
        }
      } catch (fontLoadError) {
        console.warn("[OfficeHelpers] Impossible de charger les styles de police originaux :", fontLoadError);
      }

      let originalTrackingMode = null;
      let isTrackingModified = false;

      if (trackChanges && typeof Word.ChangeTrackingMode !== "undefined") {
        try {
          context.document.load("changeTrackingMode");
          await context.sync();
          originalTrackingMode = context.document.changeTrackingMode;
          
          context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
          await context.sync();
          isTrackingModified = true;
        } catch (trackError) {
          console.warn("[OfficeHelpers] Impossible de configurer le suivi des modifications :", trackError);
        }
      }

      const insertedRange = selection.insertText(text, wordLocation);

      // Réappliquer le formatage sauvegardé sur le nouveau range
      if (insertedRange && insertedRange.font) {
        Object.keys(fontBackup).forEach(prop => {
          try {
            insertedRange.font[prop] = fontBackup[prop];
          } catch (applyError) {
            console.warn(`[OfficeHelpers] Impossible de réappliquer la propriété font.${prop} :`, applyError);
          }
        });
      }

      await context.sync();

      if (isTrackingModified && typeof Word.ChangeTrackingMode !== "undefined") {
        try {
          context.document.changeTrackingMode = originalTrackingMode;
          await context.sync();
        } catch (restoreError) {
          console.warn("[OfficeHelpers] Impossible de restaurer le suivi des modifications :", restoreError);
        }
      }
    });
  }

  // =========================================================================
  // ACTIONS EXCEL
  // =========================================================================

  /**
   * Récupère les données et formules de la sélection Excel actuelle.
   * Limite l'analyse à la zone utilisée du document pour éviter les dépassements de mémoire.
   */
  async getSelectedExcelData() {
    if (this.getHost() !== "Excel") return null;

    return Excel.run(async (context) => {
      const activeSheet = context.workbook.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      
      // On charge l'adresse brute d'abord
      range.load("address");
      await context.sync();

      // Pour éviter de lire des lignes/colonnes entières vides, on intersecte la sélection avec la zone utilisée de la feuille
      const usedRange = activeSheet.getUsedRange();
      const intersection = range.getIntersectionOrNullObject(usedRange);
      intersection.load(["values", "formulas", "address", "rowCount", "columnCount"]);
      
      await context.sync();

      if (intersection.isNullObject) {
        return {
          address: range.address,
          values: [[""]],
          formulas: [[""]],
          rowCount: 1,
          colCount: 1,
          hasData: false,
          markdown: ""
        };
      }

      const values = intersection.values;
      const formulas = intersection.formulas;
      const rowCount = intersection.rowCount;
      const colCount = intersection.columnCount;

      // Détermine si des cellules contiennent des valeurs non vides
      let hasData = false;
      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          if (values[r][c] !== null && values[r][c] !== undefined && values[r][c] !== "") {
            hasData = true;
            break;
          }
        }
      }

      // Convertir la grille 2D en tableau Markdown lisible par l'IA
      let markdown = "";
      if (hasData) {
        markdown += `Plage de cellules sélectionnée : \`${intersection.address}\` (${rowCount} ligne(s) x ${colCount} colonne(s))\n\n`;
        // En-têtes (si plusieurs lignes, on suppose que la 1ère ligne est l'en-tête)
        const headers = values[0].map((v, i) => (v !== null && v !== "" ? String(v) : `Col ${i + 1}`));
        markdown += `| ${headers.join(" | ")} |\n`;
        markdown += `| ${headers.map(() => "---").join(" | ")} |\n`;

        // Lignes suivantes
        for (let r = 1; r < values.length; r++) {
          const rowValues = values[r].map(v => (v !== null ? String(v) : ""));
          markdown += `| ${rowValues.join(" | ")} |\n`;
        }
      }

      return {
        address: intersection.address,
        values,
        formulas,
        rowCount,
        colCount,
        hasData,
        markdown
      };
    });
  }

  /**
   * Écrit une valeur ou une formule dans la cellule active ou la plage sélectionnée.
   * @param {string|Array<Array<string>>} data - Les valeurs/formules à insérer.
   * @param {boolean} isFormula - True si la valeur doit être interprétée comme une formule.
   */
  async writeExcelSelection(data, isFormula = false) {
    if (this.getHost() !== "Excel") return;

    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      
      // Assurer que la donnée est structurée en tableau 2D pour Office.js
      let grid = data;
      if (!Array.isArray(data)) {
        grid = [[data]];
      }

      // Charger les limites pour éviter les dépassements de taille de plage
      range.load(["rowCount", "columnCount"]);
      await context.sync();

      // Redimensionner le range cible pour qu'il corresponde exactement à la taille du tableau injecté
      const targetRows = grid.length;
      const targetCols = grid[0].length;
      const targetRange = range.getCell(0, 0).getResizedRange(targetRows - 1, targetCols - 1);

      if (isFormula) {
        targetRange.formulas = grid;
      } else {
        targetRange.values = grid;
      }

      targetRange.select();
      await context.sync();
    });
  }

  /**
   * Insère une image (au format base64) dans le document Word à la position de la sélection.
   * @param {string} base64Image - Données de l'image en base64 (avec ou sans préfixe).
   * @param {'replace' | 'after' | 'before'} location - Emplacement de l'insertion.
   */
  async insertImageWord(base64Image, location = "replace") {
    if (this.getHost() !== "Word") return;

    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      let wordLocation;
      
      switch (location) {
        case "after":
          wordLocation = Word.InsertLocation.after;
          break;
        case "before":
          wordLocation = Word.InsertLocation.before;
          break;
        case "replace":
        default:
          wordLocation = Word.InsertLocation.replace;
          break;
      }
      
      const rawBase64 = base64Image.replace(/^data:image\/\w+;base64,/, "");
      selection.insertInlinePictureFromBase64(rawBase64, wordLocation);
      await context.sync();
    });
  }

  /**
   * Insère une image (au format base64) dans la feuille Excel active.
   * @param {string} base64Image - Données de l'image en base64 (avec ou sans préfixe).
   */
  async insertImageExcel(base64Image) {
    if (this.getHost() !== "Excel") return;

    return Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const rawBase64 = base64Image.replace(/^data:image\/\w+;base64,/, "");
      const shape = activeSheet.shapes.addImage(rawBase64);
      
      try {
        const range = context.workbook.getSelectedRange();
        range.load(["top", "left"]);
        await context.sync();
        
        shape.left = range.left;
        shape.top = range.top;
      } catch (e) {
        console.warn("[OfficeHelpers] Impossible de positionner le shape par rapport à la sélection:", e);
      }
      
      await context.sync();
    });
  }
}

// Exposer pour les scripts
window.OfficeHelpers = OfficeHelpers;

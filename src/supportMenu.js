/**
 * @typedef {Object} DataRanges
 * @property {string} support
 * @property {string} norm
 * @property {string?} event
 */

/**
 * @typedef {Object} DefaultRowData
 * @property {Array<string>} support
 * @property {Array<string>} norm
 * @property {Array<string>?} event
 */

/**
 * @typedef {Object} BlacklistDataRanges
 * @property {string} support
 * @property {string} keys
 */

const SUPPORT_MENU_TITLE = "💟 Саппорты 💟";

class SupportMenu {
  static addMenuToTable() {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu(SUPPORT_MENU_TITLE)
      .addSubMenu(ui
        .createMenu("⌛ Норма")
        .addItem("🟰 Перевести время в дес. дробь", "SupportMenu.convertTimeToDecimal")
        .addItem("🗑️ Очистить норму", "SupportMenu.deleteSupportTime") 
      )
      .addSubMenu(ui
        .createMenu("📓 Чёрный список")
        .addItem("❔ Проверить", "SupportMenu.checkBlacklist")
      )
      .addItem("🗑️ Удалить саппорта", "SupportMenu.deleteSupport")
      .addToUi();
  }

  static convertTimeToDecimal() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getName() != "Норма") {
      Browser.msgBox(`Чтобы использовать эту функцию, вы должны выбрать диапазон на листе "Норма"`);
    }

    const selectedRange = sheet.getActiveRange();
    const data = selectedRange.getDisplayValues();

    for (let i = 0; i < data.length; i++) {
      data[i] = data[i].map(x => {
        if (x.split(" ").length !== 2) {
          return x;
        }

        const time = x.split(" ");
        const hours = Number(time[0]);
        const minutes = Number(time[1]);

        return (hours + (minutes / 60)).toFixed(2).replace(".", ",");
      });
    }

    selectedRange.setValues(data);
  }

  static deleteSupportTime() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Норма");
    const superiorSupportTime = sheet.getRange("E5:K21");
    const supportTime = sheet.getRange("E24:K73");

    /** 
     * @method
     * @returns {boolean}
     */
    const confirmDelete = () => {
      return ui.alert("❗ Вы уверены, что хотите удалить норму? ❗", ui.ButtonSet.YES_NO) === ui.Button.YES;
    } 

    if (confirmDelete()) {
      superiorSupportTime.clearContent();
      supportTime.clearContent(); 
    }
  }

  static checkBlacklist() {
    const ui = SpreadsheetApp.getUi();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ЧС");
    
    /** @type {DataRange} */
    const dataRanges = Object.create({}, {
      support: {
        value: "D5:D64"
      },
      keys: {
        value: "J5:J64"
      }
    });

    const supportBlacklist = sheet.getRange(dataRanges.support).getDisplayValues();
    const keysBlacklist = sheet.getRange(dataRanges.keys).getDisplayValues();

    /** 
     * @method
     * @returns {{ id: string, buttonResult: Button }}
     */
    const windowInteractionResult = () => {
      const window = ui.prompt(
        "Проверить саппорта в ЧС (+ Ключи)",
        "Укажите ID",
        ui.ButtonSet.OK_CANCEL
      );

      return { id: window.getResponseText(), buttonResult: window.getSelectedButton() };
    }

    /** 
     * @method
     * @arg {string} id
     * @arg {Array<string>} blacklist
     * @returns {{ appears: boolean, position: number | null }}
     */
    const tryToFindInBlacklist = (id, blacklist) => {
      const index = blacklist.findIndex(x => x.includes(id));

      return {
        appears: index != -1,
        position: (index != -1) ? index + 1 : null  
      };
    } 

    /**
     * @method
     * @arg {Array<{ name: string, value: {appears: boolean, position: number | null} }>} blacklistResults
     */
    const displayResult = (blacklistResults) => {
      let result = String();

      for (let blacklistResult of blacklistResults) {
        if (blacklistResult.value.appears) {
          result = result.concat(`ID был найден в черном списке. [Список: "${blacklistResult.name}", Номер строки: ${blacklistResult.value.position}]\n`);
        }
      }

      if (result.length <= 0) {
        ui.alert("ID не был найден в черных списках");
        return;
      }

      ui.alert(result);
    }

    const interactionResult = windowInteractionResult();

    if (interactionResult.buttonResult != ui.Button.OK) {
      return;
    }

    if (interactionResult.id.length < 17 || interactionResult.id.length > 22) {
      ui.alert("Вы ввели пустой или некорректный ID");
      return;
    }

    displayResult([
      { name: "Саппорты", value: tryToFindInBlacklist(interactionResult.id, supportBlacklist)},
      { name: "Ключи", value: tryToFindInBlacklist(interactionResult.id, keysBlacklist)}
    ]);
  }

  static deleteSupport() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadsheet.getActiveSheet();
    const normSheet = spreadsheet.getSheetByName("Норма");
    const eventSheet = spreadsheet.getSheetByName("Ивент");
    const activeCell = activeSheet.getActiveCell();

    /** @type {string} */
    const supportId = activeCell.getValue();

    /** @type {DataRanges} */
    const dataRanges = Object.create({}, {
      support: {
        value: "L5:T54"
      },
      norm: {
        value: "E24:K73"
      },
      event: {
        value: "E5:I54"
      }
    });

    /** @type {DefaultRowData} */
    const defaultRowData = Object.create({}, {
      support: {
        value: ["", "", "-", "", "<@>", "", "", "", "0/3"]
      },
      norm: {
        value: ["", "", "", "", "", "", ""]
      },
      event: {
        value: ["", "", "", "", ""]
      }
    });

    /**
     * @method
     * @arg {string} id
     * @arg {Array<Array<string>>} data
     * @returns {number}
     */
    const findSupportIndex = (id, data) => {
      let index = null;

      for (let i = 0; i < data.length; i++) {
        if (data[i][4] === id) {   // 4 ==> ID
          index = i;
          break;
        }
      }

      return index;
    } 

    /** 
     * @method
     * @arg {SpreadsheetApp.Range} range
     * @arg {Array<Array<string>>} data
     * @arg {number} index
     * @arg {Array<string>} defaultValue
     */
    const deleteAndInsertDefaultValue = (range, data, index, defaultValue) => {
      data.splice(index, 1);
      data.push(defaultValue);

      range.setValues(data);
    }

    /**
     * @method 
     * @returns {boolean} */
    const confirmDelete = () => {
      return ui.alert("❗ Вы уверены, что хотите удалить саппорта? ❗", ui.ButtonSet.YES_NO) === ui.Button.YES;
    }

    /** @method */
    const deleteFromScheduleAndNorm = () => {
      if (!confirmDelete()) return;

      const supportRange = activeSheet.getRange(dataRanges.support);
      const normRange = normSheet.getRange(dataRanges.norm);
      const eventRange = eventSheet.getRange(dataRanges.event);

      const supportData = supportRange.getDisplayValues();
      const normData = normRange.getDisplayValues();
      const eventData = eventRange.getDisplayValues();

      const index = findSupportIndex(supportId, supportData);

      deleteAndInsertDefaultValue(supportRange, supportData, index, defaultRowData.support);
      deleteAndInsertDefaultValue(normRange, normData, index, defaultRowData.norm);
      deleteAndInsertDefaultValue(eventRange, eventData, index, defaultRowData.event);
    }

    try {
      if (activeSheet.getName() != "Список") {
        throw new Error(`Чтобы удалить саппорта, вы должны находиться на листе "Список" 📃`);
      }
      
      if (activeCell.getColumn() != 16) { // Столбец ID
        throw new Error("Вы должны выбрать ID саппорта 🆔");
      }

      if (supportId.length < 17 || supportId.length > 22) {
        throw new Error("Вы выбрали саппорта с пустым или некорректным ID ❕");
      }

      deleteFromScheduleAndNorm();
    }
    catch (error) {
      ui.alert(`Произошла ошибка: ${error.message}`);
    }
  }

  static setReprimandNotes(event) {
    /** @type {SpreadsheetApp.Range} */
    const range = event.range;
    const row = range.getRow();

    if ((range.getSheet().getName() === "Список" && range.getColumn() === 20) && (row >= 5 && row <= 54)) {
      if (range.getValue() === "0/3") {
        range.setNote(null);
        return;
      }
      
      range.setNote(`Дата и время выговора: ${new Date().toLocaleString()}`);
    }
  }
}
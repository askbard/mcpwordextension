export type OfficeApp = "Word" | "Excel" | "PowerPoint" | "Unknown";

export function getOfficeApp(): OfficeApp {
  if (typeof Office === "undefined") return "Unknown";
  const host = Office.context.host;
  if (host === Office.HostType.Word) return "Word";
  if (host === Office.HostType.Excel) return "Excel";
  if (host === Office.HostType.PowerPoint) return "PowerPoint";
  return "Unknown";
}

export async function getSelectedText(): Promise<string> {
  const app = getOfficeApp();

  if (app === "Word") {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        resolve(selection.text);
      }).catch(reject);
    });
  }

  if (app === "Excel") {
    return new Promise((resolve, reject) => {
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("text");
        await context.sync();
        const values = range.text as string[][];
        resolve(values.flat().filter(Boolean).join("\t"));
      }).catch(reject);
    });
  }

  if (app === "PowerPoint") {
    return new Promise((resolve, reject) => {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value as string);
          } else {
            reject(new Error(result.error.message));
          }
        }
      );
    });
  }

  return "";
}

export async function insertTextAtCursor(text: string): Promise<void> {
  const app = getOfficeApp();

  if (app === "Word") {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  if (app === "Excel") {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      range.values = [[text]];
      await context.sync();
    });
  }

  if (app === "PowerPoint") {
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error.message));
          }
        }
      );
    });
  }
}

export async function insertTextBelow(text: string): Promise<void> {
  const app = getOfficeApp();

  if (app === "Word") {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertParagraph(text, Word.InsertLocation.after);
      await context.sync();
    });
  }

  // For Excel and PowerPoint, fall back to replace
  return insertTextAtCursor(text);
}

export async function getDocumentText(): Promise<string> {
  const app = getOfficeApp();

  if (app === "Word") {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        resolve(body.text);
      }).catch(reject);
    });
  }

  return "";
}

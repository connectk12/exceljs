type SanitizeTextOpts = {
  removeSpecialChars?: boolean;
  removeWhitespace?: boolean;
  uppercase?: boolean;
};

/**
 * Sanitizes a given text string based on the provided options.
 * It can remove special characters, remove extra whitespace, and convert the text to uppercase.
 *
 * @param {string | (string | undefined)} text - The text to be sanitized.
 * @param {SanitizeTextOpts} [opts] - Options to customize the sanitization process.
 * @param {boolean} [opts.removeSpecialChars=true] - Whether to remove special characters from the text.
 * @param {boolean} [opts.removeWhitespace=false] - Whether to remove extra whitespace from the text.
 * @param {boolean} [opts.uppercase=true] - Whether to convert the text to uppercase.
 *
 * @returns {string | (string | undefined)} - The sanitized text.
 */

export const sanitizeText = <T extends string | (string | undefined)>(
  text: T,
  opts: SanitizeTextOpts = {
    removeSpecialChars: true,
    removeWhitespace: false,
    uppercase: true,
  }
): T => {
  if (text && typeof text === "string") {
    let sanitizedText: string = text;

    // Sanitize
    if (opts.removeSpecialChars) {
      sanitizedText = sanitizedText.replace(/[^a-zA-Z0-9\s]/g, "");
    }
    if (opts.removeWhitespace) {
      sanitizedText = sanitizedText.replace(/\s+/g, " ");
    }
    if (opts.uppercase) {
      sanitizedText = sanitizedText.toLocaleUpperCase();
    }

    // Trim
    sanitizedText = sanitizedText.trim();
    return sanitizedText as T;
  }
  return undefined as T;
};

export const colLetterToNumber = (colLetter: string) => {
  let colNumber = 0;
  for (let i = 0; i < colLetter.length; i++) {
    colNumber +=
      (colLetter.charCodeAt(i) - 64) * Math.pow(26, colLetter.length - i - 1);
  }
  return colNumber;
};

export const stdTimezoneOffset = (date?: Date) => {
  var jan = new Date(new Date(date ?? "").getFullYear(), 0, 1);
  var jul = new Date(new Date(date ?? "").getFullYear(), 6, 1);
  return Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
};

export const isDstObserved = (date?: Date) => {
  return new Date(date ?? "").getTimezoneOffset() < stdTimezoneOffset(date);
};

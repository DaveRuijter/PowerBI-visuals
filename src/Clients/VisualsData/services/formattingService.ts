/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved. 
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *   
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *   
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

/// <reference path="../_references.ts"/>

module powerbi {

    import StringExtensions = jsCommon.StringExtensions;
    import Formatting = jsCommon.Formatting;

    /** Culture interfaces. These match the Globalize library interfaces intentionally. */
    export interface Culture {
        name: string;
        calendar: Calendar;
        calendars: CalendarDictionary;
        numberFormat: NumberFormatInfo;
    }

    export interface Calendar {
        patterns: any;
        firstDay: number;
    }

    export interface CalendarDictionary {
        [key: string]: Calendar;
    }

    export interface NumberFormatInfo {
        decimals: number;
        groupSizes: number[];
        negativeInfinity: string;
        positiveInfinity: string;
    }

    /** Formatting Encoder */
    module FormattingEncoder {
        export function preserveEscaped(format: string, specialChars: string): string {
            // Unicode U+E000 - U+F8FF is a private area and so we can use the chars from the range to encode the escaped sequences
            let length = specialChars.length;
            for (let i = 0; i < length; i++) {
                let oldText = "\\" + specialChars[i];
                let newText = String.fromCharCode(0xE000 + i);
                format = StringExtensions.replaceAll(format, oldText, newText);
            }
            return format;
        }

        export function restoreEscaped(format: string, specialChars: string): string {
            // After formatting is complete we should restore the encoded escaped chars into the unescaped chars
            let length = specialChars.length;
            for (let i = 0; i < length; i++) {
                let oldText = String.fromCharCode(0xE000 + i);
                let newText = specialChars[i];
                format = StringExtensions.replaceAll(format, oldText, newText);
            }
            return StringExtensions.replaceAll(format, "\\", "");
        }

        export function preserveLiterals(format: string, literals: string[]): string {
            // Unicode U+E000 - U+F8FF is a private area and so we can use the chars from the range to encode the escaped sequences
            format = StringExtensions.replaceAll(format, "\"", "'");
            for (let i = 0; ; i++) {
                let fromIndex = format.indexOf("'");
                if (fromIndex < 0) {
                    break;
                }
                let toIndex = format.indexOf("'", fromIndex + 1);
                if (toIndex < 0) {
                    break;
                }
                let literal = format.substring(fromIndex, toIndex + 1);
                literals.push(literal.substring(1, toIndex - fromIndex));
                let token = String.fromCharCode(0xE100 + i);
                format = format.replace(literal, token);
            }
            return format;
        }

        export function restoreLiterals(format: string, literals: string[]): string {
            let count = literals.length;
            for (let i = 0; i < count; i++) {
                let token = String.fromCharCode(0xE100 + i);
                let literal = literals[i];
                format = format.replace(token, literal);
            }
            return format;
        }
    }

    /** Formatting Service */ 
    class FormattingService implements IFormattingService {

        _currentCultureSelector: string;
        _currentCulture: Culture;
        _dateTimeScaleFormatInfo: DateTimeScaleFormatInfo;

        public formatValue(value: any, format?: string, culture?: string): string {
            // Handle special cases
            if (value === undefined || value === null) {
                return '';
            }
            let gculture = this.getCulture(culture);

            if (DateTimeFormat.canFormat(value)) {
                // Dates
                return DateTimeFormat.format(value, format, gculture);
            } else if (NumberFormat.canFormat(value)) {
                // Numbers
                return NumberFormat.format(value, format, gculture);
            } else {
                // Other data types - return as string
                return value.toString();
            }
        }

        public format(formatWithIndexedTokens: string, args: any[], culture?: string): string {
            if (!formatWithIndexedTokens) {
                return "";
            }
            let result = formatWithIndexedTokens.replace(/({{)|(}})|{(\d+[^}]*)}/g, (match: string, left: string, right: string, argToken: string) => {
                if (left) {
                    return "{";
                } else if (right) {
                    return "}";
                } else {
                    let parts = argToken.split(":");
                    let argIndex = parseInt(parts[0], 10);
                    let argFormat = parts[1];
                    return this.formatValue(args[argIndex], argFormat, culture);
                }
                return "";
            });

            return result;
        }

        public isStandardNumberFormat(format: string): boolean {
            return NumberFormat.isStandardFormat(format);
        }

        public formatNumberWithCustomOverride(value: number, format: string, nonScientificOverrideFormat: string, culture?: string): string {
            let gculture = this.getCulture(culture);

            return NumberFormat.formatWithCustomOverride(value, format, nonScientificOverrideFormat, gculture);
        }

        public dateFormatString(unit: DateTimeUnit): string {
            return this._dateTimeScaleFormatInfo.getFormatString(unit);
        }

        /** 
         * Sets the current localization culture
         * @param cultureSelector - name of a culture: "en", "en-UK", "fr-FR" etc. (See National Language Support (NLS) for full lists. Use "default" for invariant culture).
         */
        private setCurrentCulture(cultureSelector: string): void {
            if (this._currentCultureSelector !== cultureSelector) {
                this._currentCulture = this.getCulture(cultureSelector);
                this._currentCultureSelector = cultureSelector;
                this._dateTimeScaleFormatInfo = new DateTimeScaleFormatInfo(this._currentCulture);
            }
        }

        /** 
         * Gets the culture assotiated with the specified cultureSelector ("en", "en-US", "fr-FR" etc). 
         * @param cultureSelector - name of a culture: "en", "en-UK", "fr-FR" etc. (See National Language Support (NLS) for full lists. Use "default" for invariant culture).
         * Exposing this function for testability of unsupported cultures 
         */
        public getCulture(cultureSelector?: string): Culture {
            if (cultureSelector == null) {
                if (this._currentCulture == null) {
                    this.initialize();
                }
                return this._currentCulture;
            } else {
                let culture = Globalize.findClosestCulture(cultureSelector);
                if (!culture)
                    culture = Globalize.culture("en-US");
                return culture;
            }
        }

        /** By default the Globalization module initializes to the culture/calendar provided in the language/culture URL params */
        private initialize() {
            let cultureName = this.getUrlParam("language") || window["cultureInfo"] || window.navigator.userLanguage || window.navigator["language"] || Globalize.culture().name;
            this.setCurrentCulture(cultureName);
            let calendarName = this.getUrlParam("calendar");
            if (calendarName) {
                let culture = this._currentCulture;
                let c = culture.calendars[calendarName];
                if (c) {
                    culture.calendar = c;
                }
            }
        }

        private getUrlParam(name: string): string {
            let param = window.location.search.match(RegExp("[?&]" + name + "=([^&]*)"));
            return param ? param[1] : undefined;
        }
    }

    /**
     * DateTimeFormat module contains the static methods for formatting the DateTimes. 
     * It extends the JQuery.Globalize functionality to support complete set of .NET 
     * formatting expressions for dates.
     */
    module DateTimeFormat {

        var _currentCachedFormat: string;
        var _currentCachedProcessedFormat: string;

        /** Evaluates if the value can be formatted using the NumberFormat */
        export function canFormat(value: any) {
            let result = value instanceof Date;
            return result;
        }

        /** Formats the date using provided format and culture */
        export function format(value: Date, format: string, culture: Culture): string {
            format = format || "G";
            let isStandard = format.length === 1;
            try {
                if (isStandard) {
                    return formatDateStandard(value, format, culture);
                } else {
                    return formatDateCustom(value, format, culture);
                }
            } catch (e) {
                return formatDateStandard(value, "G", culture);
            }
        }

        /** Formats the date using standard format expression */
        function formatDateStandard(value: Date, format: string, culture: Culture) {
            // In order to provide parity with .NET we have to support additional set of DateTime patterns.
            let patterns = culture.calendar.patterns;
            // Extend supported set of patterns
            ensurePatterns(culture.calendar);
            // Handle extended set of formats
            let output = Formatting.findDateFormat(value, format, culture.name);
            if (output.format.length === 1)
                format = patterns[output.format];
            else
                format = output.format;
            //need to revisit when globalization is enabled
            culture = Globalize.culture("en-US");
            return Globalize.format(output.value, format, culture);
        }

        /** Formats the date using custom format expression */
        function formatDateCustom(value: Date, format: string, culture: Culture): string {
            let result: string;
            let literals: string[] = [];
            format = FormattingEncoder.preserveEscaped(format, "\\dfFghHKmstyz:/%'\"");
            format = FormattingEncoder.preserveLiterals(format, literals);
            format = StringExtensions.replaceAll(format, "\"", "'");
            if (format.indexOf("F") > -1) {
                // F is not supported so we need to replace the F with f based on the milliseconds                        
                // Replace all sequences of F longer than 3 with "FFF"
                format = StringExtensions.replaceAll(format, "FFFF", "FFF");
                // Based on milliseconds update the format to use fff
                let milliseconds = value.getMilliseconds();
                if (milliseconds % 10 >= 1) {
                    format = StringExtensions.replaceAll(format, "FFF", "fff");
                }
                format = StringExtensions.replaceAll(format, "FFF", "FF");
                if ((milliseconds % 100) / 10 >= 1) {
                    format = StringExtensions.replaceAll(format, "FF", "ff");
                }
                format = StringExtensions.replaceAll(format, "FF", "F");
                if ((milliseconds % 1000) / 100 >= 1) {
                    format = StringExtensions.replaceAll(format, "F", "f");
                }
                format = StringExtensions.replaceAll(format, "F", "");
                if (format === "" || format === "%")
                    return "";
            }
            format = processCustomDateTimeFormat(format);
            result = Globalize.format(value, format, culture);
            result = localize(result, culture.calendar);
            result = FormattingEncoder.restoreLiterals(result, literals);
            result = FormattingEncoder.restoreEscaped(result, "\\dfFghHKmstyz:/%'\"");
            return result;
        }

        /** Translates unsupported .NET custom format expressions to the custom expressions supported by JQuery.Globalize */
        function processCustomDateTimeFormat(format: string): string {
            if (format === _currentCachedFormat) {
                return _currentCachedProcessedFormat;
            }
            _currentCachedFormat = format;
            format = Formatting.fixDateTimeFormat(format);
            _currentCachedProcessedFormat = format;
            return format;
        }

        /** Localizes the time separator symbol */
        function localize(value: string, dictionary: any): string {
            let timeSeparator = dictionary[":"];
            if (timeSeparator === ":") {
                return value;
            }
            let result = "";
            let count = value.length;
            for (let i = 0; i < count; i++) {
                let char = value.charAt(i);
                switch (char) {
                    case ":":
                        result += timeSeparator;
                        break;
                    default:
                        result += char;
                        break;
                }
            }
            return result;
        }

        function ensurePatterns(calendar: GlobalizeCalendar) {
            let patterns = calendar.patterns;
            if (patterns["g"] === undefined) {
                patterns["g"] = patterns["f"].replace(patterns["D"], patterns["d"]);  // Generic: Short date, short time
                patterns["G"] = patterns["F"].replace(patterns["D"], patterns["d"]);  // Generic: Short date, long time
            }
        }

    }

    /** 
     * NumberFormat module contains the static methods for formatting the numbers. 
     * It extends the JQuery.Globalize functionality to support complete set of .NET 
     * formatting expressions for numeric types including custom formats.
     */
    export module NumberFormat {

        export interface NumericFormatMetadata {
            format: string;
            hasEscapes: boolean;
            hasQuotes: boolean;
            hasE: boolean;
            hasCommas: boolean;
            hasDots: boolean;
            hasPercent: boolean;
            hasPermile: boolean;
            precision: number;
            scale: number;
        }

        var _lastCustomFormatMeta: NumericFormatMetadata;

        /** Evaluates if the value can be formatted using the NumberFormat */
        export function canFormat(value: any) {
            let result = typeof (value) === "number";
            return result;
        }

        export function isStandardFormat(format: string): boolean {
            debug.assertValue(format, 'format');

            let standardFormatRegex = /^[a-z]\d{0,2}$/gi;  // a letter + up to 2 digits for precision specifier
            return standardFormatRegex.test(format);
        }

        /** Formats the number using specified format expression and culture */
        export function format(
            value: number,
            format: string,
            culture: Culture): string {
            format = format || "G";
            try {
                if (isStandardFormat(format))
                    return formatNumberStandard(value, format, culture);

                return formatNumberCustom(value, format, culture);
            } catch (e) {
                return Globalize.format(value, undefined, culture);
            }
        }

        /** Performs a custom format with a value override.  Typically used for custom formats showing scaled values. */
        export function formatWithCustomOverride(
            value: number,
            format: string,
            nonScientificOverrideFormat: string,
            culture: Culture): string {
            debug.assertValue(value, 'value');
            debug.assertValue(format, 'format');
            debug.assertValue(nonScientificOverrideFormat, 'nonScientificOverrideFormat');
            debug.assertValue(culture, 'culture');
            debug.assert(!isStandardFormat(format), 'Standard format');

            return formatNumberCustom(value, format, culture, nonScientificOverrideFormat);
        }

        /** Formats the number using standard format expression */
        function formatNumberStandard(value: number, format: string, culture: Culture): string {
            let result: string;
            let precision = <number>(format.length > 1 ? parseInt(format.substr(1, format.length - 1), 10) : undefined);
            let numberFormatInfo = culture.numberFormat;
            let formatChar = format.charAt(0);
            switch (formatChar) {
                case "e":
                case "E":
                    if (precision === undefined) {
                        precision = 6;
                    }
                    let mantissaDecimalDigits = StringExtensions.repeat("0", precision);
                    format = "0." + mantissaDecimalDigits + formatChar + "+000";
                    result = formatNumberCustom(value, format, culture);
                    break;
                case "f":
                case "F":
                    result = precision !== undefined ? value.toFixed(precision) : value.toFixed(numberFormatInfo.decimals);
                    result = localize(result, numberFormatInfo);
                    break;
                case "g":
                case "G":
                    let abs = Math.abs(value);
                    if (abs === 0 || (1E-4 <= abs && abs < 1E15)) {
                        // For the range of 0.0001 to 1,000,000,000,000,000 - use the normal form
                        result = precision !== undefined ? value.toPrecision(precision) : value.toString();
                    } else {
                        // Otherwise use exponential
                        result = precision !== undefined ? value.toExponential(precision) : value.toExponential();
                        result = result.replace("e", "E");
                    }
                    result = localize(result, numberFormatInfo);
                    break;
                case "r":
                case "R":
                    result = value.toString();
                    result = localize(result, numberFormatInfo);
                    break;
                case "x":
                case "X":
                    result = value.toString(16);
                    if (formatChar === "X") {
                        result = result.toUpperCase();
                    }
                    if (precision !== undefined) {
                        let actualPrecision = result.length;
                        let isNegative = value < 0;
                        if (isNegative) {
                            actualPrecision--;
                        }
                        let paddingZerosCount = precision - actualPrecision;
                        let paddingZeros = undefined;
                        if (paddingZerosCount > 0) {
                            paddingZeros = StringExtensions.repeat("0", paddingZerosCount);
                        }
                        if (isNegative) {
                            result = "-" + paddingZeros + result.substr(1);
                        } else {
                            result = paddingZeros + result;
                        }
                    }
                    result = localize(result, numberFormatInfo);
                    break;
                default:
                    result = Globalize.format(value, format, culture);
            }
            return result;
        }

        /** Formats the number using custom format expression */
        function formatNumberCustom(
            value: number,
            format: string,
            culture: Culture,
            nonScientificOverrideFormat?: string): string {
            let result: string;
            let numberFormatInfo = culture.numberFormat;
            if (isFinite(value)) {
                // Split format into positive;negative;zero patterns
                let signSpecificFormats = format.split(";");
                if (signSpecificFormats.length > 1) {
                    let negativeFormat = format;
                    let positiveFormat = format;
                    let zeroFormat = format;
                    if (signSpecificFormats.length === 2) {
                        positiveFormat = zeroFormat = signSpecificFormats[0];
                        negativeFormat = signSpecificFormats[1];
                    } else {
                        positiveFormat = signSpecificFormats[0];
                        negativeFormat = signSpecificFormats[1];
                        zeroFormat = signSpecificFormats[2];
                    }
                    // Pick a format based on the sign of value
                    if (value > 0) {
                        format = positiveFormat;
                    } else if (value === 0) {
                        format = zeroFormat;
                    } else {
                        format = negativeFormat;
                    }
                    value = Math.abs(value);
                }

                // Get format metadata
                let formatMeta = getCustomFormatMetadata(format);

                // Preserve literals and escaped chars
                if (formatMeta.hasEscapes) {
                    format = FormattingEncoder.preserveEscaped(format, "\\0#.,%‰");
                }
                let literals: string[] = [];
                if (formatMeta.hasQuotes) {
                    format = FormattingEncoder.preserveLiterals(format, literals);
                }

                // Scientific format
                if (formatMeta.hasE && !nonScientificOverrideFormat) {
                    let scientificMatch = /e[+-]*0+/gi.exec(format);
                    if (scientificMatch) {
                        // Case 2.1. Scientific custom format
                        let formatM = format.substr(0, scientificMatch.index);
                        let formatE = format.substr(scientificMatch.index + scientificMatch[0].indexOf("0"));
                        let precision = getCustomFormatPrecision(formatM, formatMeta);
                        let scale = getCustomFormatScale(formatM, formatMeta);
                        if (scale !== 1) {
                            value = value * scale;
                        }
                        let s = value.toExponential(precision);
                        let indexOfE = s.indexOf("e");
                        let mantissa = s.substr(0, indexOfE);
                        let exp = s.substr(indexOfE + 1);
                        let resultM = fuseNumberWithCustomFormat(mantissa, formatM, numberFormatInfo);
                        let resultE = fuseNumberWithCustomFormat(exp, formatE, numberFormatInfo);
                        if (resultE.charAt(0) === "+" && scientificMatch[0].charAt(1) !== "+") {
                            resultE = resultE.substr(1);
                        }
                        let e = scientificMatch[0].charAt(0);
                        result = resultM + e + resultE;
                    }
                }

                // Non scientific format
                if (result === undefined) {
                    let valueFormatted: string;
                    if (nonScientificOverrideFormat) {
                        valueFormatted = formattingService.format(nonScientificOverrideFormat, [value], culture.name);
                    } else {
                        let precision = getCustomFormatPrecision(format, formatMeta);
                        let scale = getCustomFormatScale(format, formatMeta);
                        if (scale !== 1) {
                            value = value * scale;
                        }
                        valueFormatted = toNonScientific(value, precision);
                    }
                    result = fuseNumberWithCustomFormat(valueFormatted, format, numberFormatInfo, !!nonScientificOverrideFormat);
                }
                if (formatMeta.hasQuotes) {
                    result = FormattingEncoder.restoreLiterals(result, literals);
                }
                if (formatMeta.hasEscapes) {
                    result = FormattingEncoder.restoreEscaped(result, "\\0#.,%‰");
                }

                _lastCustomFormatMeta = formatMeta;
            } else {
                return Globalize.format(value, undefined);
            }
            return result;
        }

        /** Returns string with the fixed point respresentation of the number */
        function toNonScientific(value: number, precision: number): string {
            let result = "";
            let precisionZeros = 0;
            // Double precision numbers support actual 15-16 decimal digits of precision.
            if (precision > 16) {
                precisionZeros = precision - 16;
                precision = 16;
            }
            let digitsBeforeDecimalPoint = Double.log10(Math.abs(value));
            if (digitsBeforeDecimalPoint < 16) {
                if (digitsBeforeDecimalPoint > 0) {
                    let maxPrecision = 16 - digitsBeforeDecimalPoint;
                    if (precision > maxPrecision) {
                        precisionZeros += precision - maxPrecision;
                        precision = maxPrecision;
                    }
                }
                result = value.toFixed(precision);
            } else if (digitsBeforeDecimalPoint === 16) {
                result = value.toFixed(0);
                precisionZeros += precision;
                if (precisionZeros > 0) {
                    result += ".";
                }
            } else { // digitsBeforeDecimalPoint > 16
                // Different browsers have different implementations of the toFixed(). 
                // In IE it returns fixed format no matter what's the number. In FF and Chrome the method returns exponential format for numbers greater than 1E21.
                // So we need to check for range and convert the to exponential with the max precision. 
                // Then we convert exponential string to fixed by removing the dot and padding with "power" zeros.
                result = value.toExponential(15);
                let indexOfE = result.indexOf("e");
                if (indexOfE > 0) {
                    let indexOfDot = result.indexOf(".");
                    let mantissa = result.substr(0, indexOfE);
                    let exp = result.substr(indexOfE + 1);
                    let powerZeros = parseInt(exp, 10) - (mantissa.length - indexOfDot - 1);
                    result = mantissa.replace(".", "") + StringExtensions.repeat("0", powerZeros);
                    if (precision > 0) {
                        result = result + "." + StringExtensions.repeat("0", precision);
                    }
                }
            }
            if (precisionZeros > 0) {
                result = result + StringExtensions.repeat("0", precisionZeros);
            }
            return result;
        }

        /** Returns the formatMetadata of the format */
        export function getCustomFormatMetadata(format: string, calculatePrecision?: boolean, calculateScale?: boolean): NumericFormatMetadata {
            if (_lastCustomFormatMeta !== undefined && format === _lastCustomFormatMeta.format) {
                return _lastCustomFormatMeta;
            }
            let result = {
                format: format,
                hasEscapes: false,
                hasQuotes: false,
                hasE: false,
                hasCommas: false,
                hasDots: false,
                hasPercent: false,
                hasPermile: false,
                precision: -1,
                scale: -1,
            };
            let length = format.length;
            for (let i = 0; i < length; i++) {
                let c = format.charAt(i);
                switch (c) {
                    case "\\":
                        result.hasEscapes = true;
                        break;
                    case "'":
                    case "\"":
                        result.hasQuotes = true;
                        break;
                    case "e":
                    case "E":
                        result.hasE = true;
                        break;
                    case ",":
                        result.hasCommas = true;
                        break;
                    case ".":
                        result.hasDots = true;
                        break;
                    case "%":
                        result.hasPercent = true;
                        break;
                    case "‰":
                        result.hasPermile = true;
                        break;
                }
            }

            if (calculatePrecision)
                result.precision = getCustomFormatPrecision(format, result);

            if (calculateScale)
                result.scale = getCustomFormatScale(format, result);

            return result;
        }
    
        /** Returns the decimal precision of format based on the number of # and 0 chars after the decimal point */
        function getCustomFormatPrecision(format: string, formatMeta: NumericFormatMetadata): number {
            if (formatMeta.precision > -1) {
                return formatMeta.precision;
            }
            let result = 0;
            if (formatMeta.hasDots) {
                let dotIndex = format.indexOf(".");
                if (dotIndex > -1) {
                    let count = format.length;
                    for (let i = dotIndex; i < count; i++) {
                        let char = format.charAt(i);
                        if (char === "#" || char === "0")
                            result++;
                    }
                    result = Math.min(19, result);
                }
            }
            formatMeta.precision = result;
            return result;
        }

        /** Returns the scale factor of the format based on the "%" and scaling "," chars in the format */
        function getCustomFormatScale(format: string, formatMeta: NumericFormatMetadata): number {
            if (formatMeta.scale > -1) {
                return formatMeta.scale;
            }
            let result = 1;
            if (formatMeta.hasPercent && format.indexOf("%") > -1) {
                result = result * 100;
            }
            if (formatMeta.hasPermile && format.indexOf("‰") > -1) {
                result = result * 1000;
            }
            if (formatMeta.hasCommas) {
                let dotIndex = format.indexOf(".");
                if (dotIndex === -1) {
                    dotIndex = format.length;
                }
                for (let i = dotIndex - 1; i > -1; i--) {
                    let char = format.charAt(i);
                    if (char === ",") {
                        result = result / 1000;
                    } else {
                        break;
                    }
                }
            }
            formatMeta.scale = result;
            return result;
        }

        function fuseNumberWithCustomFormat(value: string, format: string, numberFormatInfo: GlobalizeNumberFormat, suppressModifyValue?: boolean): string {
            let formatParts = format.split(".", 2);
            if (formatParts.length === 2) {
                let wholeFormat = formatParts[0];
                let fractionFormat = formatParts[1];

                let valueParts = value.split(".", 2);
                let wholeValue = valueParts[0];
                let fractionValue = valueParts.length === 2 ? valueParts[1].replace(/0+$/, "") : "";

                let wholeFormattedValue = fuseNumberWithCustomFormatLeft(wholeValue, wholeFormat, numberFormatInfo, suppressModifyValue);
                let fractionFormattedValue = fuseNumberWithCustomFormatRight(fractionValue, fractionFormat, suppressModifyValue);

                if (fractionFormattedValue.fmtOnly || fractionFormattedValue.value === "")
                    return wholeFormattedValue + fractionFormattedValue.value;

                return wholeFormattedValue + numberFormatInfo["."] + fractionFormattedValue.value;
            }
            return fuseNumberWithCustomFormatLeft(value, format, numberFormatInfo, suppressModifyValue);
        }

        function fuseNumberWithCustomFormatLeft(value: string, format: string, numberFormatInfo: GlobalizeNumberFormat, suppressModifyValue?: boolean): string {
            let groupSymbolIndex = format.indexOf(",");
            let enableGroups = groupSymbolIndex > -1 && groupSymbolIndex < Math.max(format.lastIndexOf("0"), format.lastIndexOf("#")) && numberFormatInfo[","];
            let groupDigitCount = 0;
            let groupIndex = 0;
            let groupSizes = numberFormatInfo.groupSizes || [3];
            let groupSize = groupSizes[0];
            let groupSeparator = numberFormatInfo[","];
            let sign = "";
            let firstChar = value.charAt(0);
            if (firstChar === "+" || firstChar === "-") {
                sign = numberFormatInfo[firstChar];
                value = value.substr(1);
            }
            let isZero = value === "0";
            let result = "";
            let leftBuffer = "";
            let vi = value.length - 1;
            let fmtOnly = true;
            // Iterate through format chars and replace 0 and # with the digits from the value string
            for (let fi = format.length - 1; fi > -1; fi--) {
                let formatChar = format.charAt(fi);
                switch (formatChar) {
                    case "0":
                    case "#":
                        fmtOnly = false;
                        if (leftBuffer !== "") {
                            result = leftBuffer + result;
                            leftBuffer = "";
                        }
                        if (!suppressModifyValue) {
                            if (vi > -1 || formatChar === "0") {
                                if (enableGroups) {
                                    // If the groups are enabled we'll need to keep track of the current group index and periodically insert group separator,
                                    if (groupDigitCount === groupSize) {
                                        result = groupSeparator + result;
                                        groupIndex++;
                                        if (groupIndex < groupSizes.length) {
                                            groupSize = groupSizes[groupIndex];
                                        }
                                        groupDigitCount = 1;
                                    } else {
                                        groupDigitCount++;
                                    }
                                }
                            }
                            if (vi > -1) {
                                if (isZero && formatChar === "#") {
                                    // Special case - if we need to format a zero value and the # symbol is used - we don't copy it into the result)
                                } else {
                                    result = value.charAt(vi) + result;
                                }
                                vi--;
                            } else if (formatChar !== "#") {
                                result = formatChar + result;
                            }
                        }
                        break;
                    case ",":
                        // We should skip all the , chars
                        break;
                    default:
                        leftBuffer = formatChar + leftBuffer;
                        break;
                }
            }

            // If the value didn't fit into the number of zeros provided in the format then we should insert the missing part of the value into the result
            if (!suppressModifyValue) {
                if (vi > -1 && result !== "") {
                    if (enableGroups) {
                        while (vi > -1) {
                            if (groupDigitCount === groupSize) {
                                result = groupSeparator + result;
                                groupIndex++;
                                if (groupIndex < groupSizes.length) {
                                    groupSize = groupSizes[groupIndex];
                                }
                                groupDigitCount = 1;
                            } else {
                                groupDigitCount++;
                            }
                            result = value.charAt(vi) + result;
                            vi--;
                        }
                    } else {
                        result = value.substr(0, vi + 1) + result;
                    }
                }
                // Insert sign in front of the leftBuffer and result
                return sign + leftBuffer + result;
            }

            if (fmtOnly)
                // If the format doesn't specify any digits to be displayed, then just return the format we've parsed up until now.
                return sign + leftBuffer + result;

            return sign + leftBuffer + value + result;
        }

        function fuseNumberWithCustomFormatRight(value: string, format: string, suppressModifyValue?: boolean): { value: string; fmtOnly?: boolean } {
            let vi = 0;
            let fCount = format.length;
            let vCount = value.length;
            if (suppressModifyValue) {
                debug.assert(fCount > 0, "Empty formatting string");

                if ((format.charAt(fCount - 1) !== "0") && (format.charAt(fCount - 1) !== "#"))
                    return {
                        value: value + format.charAt(fCount - 1),
                        fmtOnly: value === "",
                    };

                return {
                    value: value,
                    fmtOnly: value === "",
                };
            }

            let result = "",
                fmtOnly: boolean = true;
            for (let fi = 0; fi < fCount; fi++) {
                let formatChar = format.charAt(fi);
                if (vi < vCount) {
                    switch (formatChar) {
                        case "0":
                        case "#":
                            result += value[vi++];
                            fmtOnly = false;
                            break;
                        default:
                            result += formatChar;
                    }
                } else {
                    if (formatChar !== "#") {
                        result += formatChar;
                        fmtOnly = fmtOnly && (formatChar !== "0");
                    }
                }
            }

            return {
                value: result,
                fmtOnly: fmtOnly,
            };
        }

        function localize(value: string, dictionary: any): string {
            let plus = dictionary["+"];
            let minus = dictionary["-"];
            let dot = dictionary["."];
            let comma = dictionary[","];
            if (plus === "+" && minus === "-" && dot === "." && comma === ",") {
                return value;
            }
            let count = value.length;
            let result = "";
            for (let i = 0; i < count; i++) {
                let char = value.charAt(i);
                switch (char) {
                    case "+":
                        result = result + plus;
                        break;
                    case "-":
                        result = result + minus;
                        break;
                    case ".":
                        result = result + dot;
                        break;
                    case ",":
                        result = result + comma;
                        break;
                    default:
                        result = result + char;
                        break;
                }
            }
            return result;
        }

    }

    /** DateTimeScaleFormatInfo is used to calculate and keep the Date formats used for different units supported by the DateTimeScaleModel */
    class DateTimeScaleFormatInfo {

        // Fields
        public YearPattern: string;
        public MonthPattern: string;
        public DayPattern: string;
        public HourPattern: string;
        public MinutePattern: string;
        public SecondPattern: string;
        public MillisecondPattern: string;

        // Constructor
        /** 
         * Creates new instance of the DateTimeScaleFormatInfo class.
         * @param culture - culture which calendar info is going to be used to derive the formats.
         */
        constructor(culture: Culture) {
            let calendar: Calendar = culture.calendar;
            let patterns: any = calendar.patterns;
            let monthAbbreviations: any = calendar["months"]["namesAbbr"];
            let cultureHasMonthAbbr: boolean = monthAbbreviations && monthAbbreviations[0];
            let yearMonthPattern: string = patterns["Y"];
            let monthDayPattern: string = patterns["M"];
            let fullPattern: string = patterns["f"];
            let longTimePattern: string = patterns["T"];
            let shortTimePattern: string = patterns["t"];
            let separator: string = fullPattern.indexOf(",") > -1 ? ", " : " ";

            let hasYearSymbol: boolean = yearMonthPattern.indexOf("yyyy'") === 0 && yearMonthPattern.length > 6 && yearMonthPattern[6] === '\'';
            this.YearPattern = hasYearSymbol ? yearMonthPattern.substr(0, 7) : "yyyy";

            let yearPos: number = fullPattern.indexOf("yy");
            let monthPos: number = fullPattern.indexOf("MMMM");
            this.MonthPattern = cultureHasMonthAbbr && monthPos > -1 ? (yearPos > monthPos ? "MMM yyyy" : "yyyy MMM") : yearMonthPattern;

            this.DayPattern = cultureHasMonthAbbr ? monthDayPattern.replace("MMMM", "MMM") : monthDayPattern;

            let minutePos: number = fullPattern.indexOf("mm");
            let pmPos: number = fullPattern.indexOf("tt");
            let shortHourPattern: string = pmPos > -1 ? shortTimePattern.replace(":mm ", "") : shortTimePattern;
            this.HourPattern = yearPos < minutePos ? this.DayPattern + separator + shortHourPattern : shortHourPattern + separator + this.DayPattern;

            this.MinutePattern = shortTimePattern;

            this.SecondPattern = longTimePattern;

            this.MillisecondPattern = longTimePattern.replace("ss", "ss.fff");

            // Special cases
            switch (culture.name) {
                case "fi-FI":
                    this.DayPattern = this.DayPattern.replace("'ta'", ""); // Fix for finish 'ta' suffix for month names.
                    this.HourPattern = this.HourPattern.replace("'ta'", "");
                    break;
            }
        }

        // Methods

        /** 
         * Returns the format string of the provided DateTimeUnit.
         * @param unit - date or time unit
         */
        public getFormatString(unit: DateTimeUnit): string {
            switch (unit) {
                case DateTimeUnit.Year:
                    return this.YearPattern;
                case DateTimeUnit.Month:
                    return this.MonthPattern;
                case DateTimeUnit.Week:
                case DateTimeUnit.Day:
                    return this.DayPattern;
                case DateTimeUnit.Hour:
                    return this.HourPattern;
                case DateTimeUnit.Minute:
                    return this.MinutePattern;
                case DateTimeUnit.Second:
                    return this.SecondPattern;
                case DateTimeUnit.Millisecond:
                    return this.MillisecondPattern;
            }

            debug.assertFail('Unexpected unit: ' + unit);
        }
    }

    export var formattingService: IFormattingService = new FormattingService();
}


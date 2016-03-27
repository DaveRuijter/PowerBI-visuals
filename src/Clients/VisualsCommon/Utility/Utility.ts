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

/**
 * Defined in host.
 */
declare var clusterUri: string;

module jsCommon {

    /**
     * Http Status code we are interested.
     */
    export enum HttpStatusCode {
        OK = 200,
        BadRequest = 400,
        Unauthorized = 401,
        Forbidden = 403,
        RequestEntityTooLarge = 413,
    }

    /**
     * Other HTTP Constants.
     */
    export module HttpConstants {
        export const ApplicationOctetStream = 'application/octet-stream';
        export const MultiPartFormData = 'multipart/form-data';
    }

    /**
     * Extensions to String class.
     */
    export module StringExtensions {
        export function format(...args: string[]) {
            let s = args[0];

            if (isNullOrUndefinedOrWhiteSpaceString(s))
                return s;

            for (let i = 0; i < args.length - 1; i++) {
                let reg = new RegExp("\\{" + i + "\\}", "gm");
                s = s.replace(reg, args[i + 1]);
            }
            return s;
        }

        /**
         * Compares two strings for equality, ignoring case.
         */
        export function equalIgnoreCase(a: string, b: string): boolean {
            return StringExtensions.normalizeCase(a) === StringExtensions.normalizeCase(b);
        }

        export function startsWithIgnoreCase(a: string, b: string): boolean {
            let normalizedSearchString = StringExtensions.normalizeCase(b);
            return StringExtensions.normalizeCase(a).indexOf(normalizedSearchString) === 0;
        }

        /** Determines whether a string contains a specified substring (while ignoring case). */
        export function containsIgnoreCase(source: string, substring: string): boolean {
            if (source == null)
                return false;

            return source.toLowerCase().indexOf(substring.toLowerCase().toString()) !== -1;
        }

        /** 
         * Normalizes case for a string.
         * Used by equalIgnoreCase method. 
         */
        export function normalizeCase(value: string): string {
            Utility.throwIfNullOrUndefined(value, StringExtensions, 'normalizeCase', 'value');

            return value.toUpperCase();
        }

        /** 
         * Is string null or empty or undefined?
         * @return True if the value is null or undefined or empty string,
         * otherwise false.
         */
        export function isNullOrEmpty(value: string): boolean {
            return (value == null) || (value.length === 0);
        }

        /** 
         * Returns true if the string is null, undefined, empty, or only includes white spaces.
         * @return True if the str is null, undefined, empty, or only includes white spaces,
         * otherwise false.
         */
        export function isNullOrUndefinedOrWhiteSpaceString(str: string): boolean {
            return StringExtensions.isNullOrEmpty(str) || StringExtensions.isNullOrEmpty(str.trim());
        }

        /**
         * Returns a value indicating whether the str contains any whitespace.
         */
        export function containsWhitespace(str: string): boolean {
            Utility.throwIfNullOrUndefined(str, this, 'containsWhitespace', 'str');

            let expr: RegExp = /\s/;
            return expr.test(str);
        }

        /**
         * Returns a value indicating whether the str is a whitespace string.
         */
        export function isWhitespace(str: string): boolean {
            Utility.throwIfNullOrUndefined(str, this, 'isWhitespace', 'str');

            return str.trim() === '';
        }

        /** 
         * Returns the string with any trailing whitespace from str removed.
         */
        export function trimTrailingWhitespace(str: string): string {
            Utility.throwIfNullOrUndefined(str, this, 'trimTrailingWhitespace', 'str');
            return str.replace(/\s+$/, '');
        }

        /**
         * Returns the string with any leading and trailing whitespace from str removed.
         */
        export function trimWhitespace(str: string): string {
            Utility.throwIfNullOrUndefined(str, this, 'trimWhitespace', 'str');
            return str.replace(/^\s+/, '').replace(/\s+$/, '');
        }

        /** 
         * Returns length difference between the two provided strings.
         */
        export function getLengthDifference(left: string, right: string) {
            Utility.throwIfNullOrUndefined(left, this, 'getLengthDifference', 'left');
            Utility.throwIfNullOrUndefined(right, this, 'getLengthDifference', 'right');

            return Math.abs(left.length - right.length);
        }

        /**
         * Repeat char or string several times.
         * @param char The string to repeat.
         * @param count How many times to repeat the string.
         */
        export function repeat(char: string, count: number): string {
            let result = "";
            for (let i = 0; i < count; i++) {
                result += char;
            }
            return result;
        }

        /**
         * Replace all the occurrences of the textToFind in the text with the textToReplace.
         * @param text The original string.
         * @param textToFind Text to find in the original string.
         * @param textToReplace New text replacing the textToFind.
         */
        export function replaceAll(text: string, textToFind: string, textToReplace: string): string {
            if (!textToFind)
                return text;

            let pattern = escapeStringForRegex(textToFind);
            return text.replace(new RegExp(pattern, 'gi'), textToReplace);
        }

        /**
         * Returns a name that is not specified in the values.
         */
        export function findUniqueName(
            usedNames: { [name: string]: boolean },
            baseName: string): string {
            debug.assertValue(usedNames, 'usedNames');
            debug.assertValue(baseName, 'baseName');

            // Find a unique name
            let i = 0,
                uniqueName: string = baseName;
            while (usedNames[uniqueName]) {
                uniqueName = baseName + (++i);
            }

            return uniqueName;
        }

        export function constructCommaSeparatedList(list: string[], resourceProvider: IStringResourceProvider, maxValue?: number): string {
            if (!list || list.length === 0)
                return '';

            if (maxValue === null || maxValue === undefined)
                maxValue = Number.MAX_VALUE;

            let length = Math.min(maxValue, list.length);

            let replacedList = [];
            // Only need to replace user entries of {0} and {1} since we build the list in pairs.
            for (let j = 0; j < 2; j++) {
                let targetValue = '{' + j + '}';
                let replaceValue = '_|_<' + j + '>_|_';
                for (let i = 0; i < length; i++) {
                    if (list[i].indexOf(targetValue) > -1) {
                        list[i] = list[i].replace(targetValue, replaceValue);
                        replacedList.push({ targetValue: targetValue, replaceValue: replaceValue });
                    }
                }
            }

            let commaSeparatedList: string = '';
            for (let i = 0; i < length; i++) {
                if (i === 0)
                    commaSeparatedList = list[i];
                else
                    commaSeparatedList = StringExtensions.format(resourceProvider.get('FilterRestatement_Comma'), commaSeparatedList, list[i]);
            }

            for (let i = 0; i < replacedList.length; i++) {
                commaSeparatedList = commaSeparatedList.replace(replacedList[i].replaceValue, replacedList[i].targetValue);
            }

            return commaSeparatedList;
        }

        export function escapeStringForRegex(s: string): string {
            return s.replace(/([-()\[\]{}+?*.$\^|,:#<!\\])/g, '\\$1');
        }
    }

    /**
     * Interface used for interacting with WCF typed objects.
     */
    export interface TypedObject {
        __type: string;
    }

    /** 
     * The general utility class.
     */
    export class Utility {
        private static TypeNamespace = 'http://schemas.microsoft.com/sqlbi/2013/01/NLRuntimeService';

        public static JsonContentType = 'application/json';
        public static JpegContentType = 'image/jpeg';
        public static XJavascriptContentType = 'application/x-javascript';
        public static JsonDataType = 'json';
        public static BlobDataType = 'blob';
        public static HttpGetMethod = 'GET';
        public static HttpPostMethod = 'POST';
        public static HttpPutMethod = 'PUT';
        public static HttpDeleteMethod = 'DELETE';
        public static HttpContentTypeHeader = 'Content-Type';
        public static HttpAcceptHeader = 'Accept';
        public static Undefined = 'undefined';

        private static staticContentLocation: string = window.location.protocol + '//' + window.location.host;

        /**
         * Ensures the specified value is not null or undefined. Throws a relevent exception if it is.
         * @param value The value to check.
         * @param context The context from which the check originated.
         * @param methodName The name of the method that initiated the check.
         * @param parameterName The parameter name of the value to check.
         */
        public static throwIfNullOrUndefined(value, context, methodName, parameterName) {
            if (value === null) {
                Utility.throwException(Errors.argumentNull(Utility.getComponentName(context) + methodName + '.' + parameterName));
            }
            else if (typeof (value) === Utility.Undefined) {
                Utility.throwException(Errors.argumentUndefined(Utility.getComponentName(context) + methodName + '.' + parameterName));
            }
        }

        /**
         * Ensures the specified value is not null, undefined or empty. Throws a relevent exception if it is.
         * @param value The value to check.
         * @param context The context from which the check originated.
         * @param methodName The name of the method that initiated the check.
         * @param parameterName The parameter name of the value to check.
         */
        public static throwIfNullOrEmpty(value: any, context: any, methodName: string, parameterName: string) {
            Utility.throwIfNullOrUndefined(value, context, methodName, parameterName);
            if (!value.length) {
                Utility.throwException(Errors.argumentOutOfRange(Utility.getComponentName(context) + methodName + '.' + parameterName));
            }
        }

        /**
         * Ensures the specified string is not null, undefined or empty. Throws a relevent exception if it is.
         * @param value The value to check.
         * @param context The context from which the check originated.
         * @param methodName The name of the method that initiated the check.
         * @param parameterName The parameter name of the value to check.
         */
        public static throwIfNullOrEmptyString(value: string, context: any, methodName: string, parameterName: string) {
            Utility.throwIfNullOrUndefined(value, context, methodName, parameterName);
            if (value.length < 1) {
                Utility.throwException(Errors.argumentOutOfRange(Utility.getComponentName(context) + methodName + '.' + parameterName));
            }
        }

        /**
         * Ensures the specified value is not null, undefined, whitespace or empty. Throws a relevent exception if it is.
         * @param value The value to check.
         * @param context The context from which the check originated.
         * @param methodName The name of the method that initiated the check.
         * @param parameterName The parameter name of the value to check.
         */
        public static throwIfNullEmptyOrWhitespaceString(value: string, context: any, methodName: string, parameterName: string) {
            Utility.throwIfNullOrUndefined(value, context, methodName, parameterName);
            if (StringExtensions.isNullOrUndefinedOrWhiteSpaceString(value)) {
                Utility.throwException(Errors.argumentOutOfRange(Utility.getComponentName(context) + methodName + '.' + parameterName));
            }
        }

        /**
         * Ensures the specified condition is true. Throws relevant exception if it isn't.
         * @param condition The condition to check.
         * @param context The context from which the check originated.
         * @param methodName The name of the method that initiated the check.
         * @param parameterName The parameter name against which the condition is checked.
         */
        public static throwIfNotTrue(condition: boolean, context: any, methodName: string, parameterName: string) {
            if (!condition) {
                Utility.throwException(Errors.argument(parameterName, Utility.getComponentName(context) + methodName + '.' + parameterName));
            }
        }

        /**
         * Checks whether the provided value is a 'string'.
         * @param value The value to test.
         */
        public static isString(value: any): boolean {
            return ((typeof value) === 'string');
        }

        /**
         * Checks whether the provided value is a 'boolean'.
         * @param value The value to test.
         */
        public static isBoolean(value: any): boolean {
            return ((typeof value) === 'boolean');
        }

        /**
         * Checks whether the provided value is a 'number'.
         * @param value The value to test.
         */
        public static isNumber(value: any): boolean {
            return ((typeof value) === 'number');
        }

        /**
         * Checks whether the provided value is a Date instance.
         * @param value The value to test.
         */
        public static isDate(value: any): boolean {
            return Utility.isObject(value) && (value instanceof Date);
        }

        /**
         * Checks whether the provided value is an 'object'.
         * @param value The value to test.
         */
        public static isObject(value: any): boolean {
            return (value != null) && ((typeof value) === 'object');
        }

        /**
         * Checks whether the provided value is null or undefined.
         * @param value The value to test.
         */
        public static isNullOrUndefined(value: any): boolean {
            return (value === null) || (typeof (value) === Utility.Undefined);
        }

        /**
         * Combine a base url and a path.
         * @param baseUrl The base url.
         * @param path The path to add on to the base url.
         * @returns The combined url.
         */
        public static urlCombine(baseUrl: string, path: string) {
            Utility.throwIfNullOrUndefined(baseUrl, null, "urlCombine", "baseUrl");
            Utility.throwIfNullOrUndefined(path, null, "urlCombine", "path");

            // should any of the components be empty, fail gracefuly - this is important when using the test page
            if (StringExtensions.isNullOrUndefinedOrWhiteSpaceString(path)) {
                return baseUrl;
            }

            if (StringExtensions.isNullOrUndefinedOrWhiteSpaceString(baseUrl)) {
                return path;
            }

            let finalUrl = baseUrl;

            if (finalUrl.charAt(finalUrl.length - 1) === '/') {
                if (path.charAt(0) === '/')
                    path = path.slice(1);
            } else {
                if (path.charAt(0) !== '/')
                    path = '/' + path;
            }

            return finalUrl + path;
        }

        public static getAbsoluteUri(path: string): string {
            Utility.throwIfNullOrUndefined(path, null, "getAbsoluteUri", "path");

            let url = path;
            // Make absolute
            if (url && url.indexOf('http') === - 1) {
                url = Utility.urlCombine(clusterUri, url);
            }
            return url;
        }

        public static getStaticResourceUri(path: string) {
            Utility.throwIfNullOrUndefined(path, null, "getStaticResourceUri", "path");

            let url = path;
            // Make absolute
            if (url && url.indexOf('http') === - 1) {
                url = jsCommon.Utility.urlCombine(Utility.staticContentLocation, url);
            }
            return url;
        }

        public static getComponentName(context) {
            return !context ? '' : (typeof context).toString() + '.';
        }

        public static throwException(e) {
            Trace.error(
                StringExtensions.format("Throwing exception: {0}", JSON.stringify(e)), 
                /*includeStackTrace*/ e.stack != null ? false : true);
            throw e;
        }

        public static createClassSelector(className: string): string {
            Utility.throwIfNullOrEmptyString(className, null, 'CreateClassSelector', 'className');
            return '.' + className;
        }

        public static createIdSelector(id: string): string {
            Utility.throwIfNullOrEmptyString(id, null, 'CreateIdSelector', 'id');
            return '#' + id;
        }

        /**
         * Creates a client-side Guid string.
         * @returns A string representation of a Guid.
         */
        public static generateGuid(): string {
            let guid = "",
                idx = 0;

            for (idx = 0; idx < 32; idx += 1) {
                let guidDigitsItem = Math.random() * 16 | 0;
                switch (idx) {
                    case 8:
                    case 12:
                    case 16:
                    case 20:
                        guid += "-";
                        break;
                }
                guid += guidDigitsItem.toString(16);
            }

            return guid;
        }

        /**
         * Generates a random 7 character string that is used as a connection group name.
         * @returns A random connection group name.
         */
        public static generateConnectionGroupName(): string {
            let name = "";
            let possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

            for (let i = 0; i < 7; i++)
                name += possible.charAt(Math.floor(Math.random() * possible.length));

            return name;
        }

        /**
         * Try extract a cookie from {@link document.cookie} identified by key.
         */
        public static getCookieValue(key: string): string {
            // the cookie is of the format <key1=value1>; <key2=value2>. Split by ';', then by '=' 
            // to search for the key
            let keyValuePairs = document.cookie.split(';');
            for (let i = 0; i < keyValuePairs.length; i++) {
                let keyValue = keyValuePairs[i];
                let split = keyValue.split('=');
                if (split.length > 0 && split[0].trim() === key) {
                    return keyValue.substr(keyValue.indexOf('=') + 1);
                }
            }
            return null;
        }

        /**
         * Extracts the protocol://hostname section of a url.
         * @param url The URL from which to extract the section.
         * @returns The protocol://hostname portion of the given URL.
         */
        public static getDomainForUrl(url: string): string {
            let hrefObject = Utility.getHrefObjectFromUrl(url);
            return hrefObject.prop('protocol') + '//' + hrefObject.prop('hostname');
        }

        /**
         * Extracts the hostname and absolute path sections of a url.
         * @param url The URL from which to extract the section.
         * @returns The hostname and absolute path portion of the given URL.
         */
        public static getHostNameForUrl(url: string): string {
            let hrefObject = Utility.getHrefObjectFromUrl(url);
            return Utility.urlCombine(hrefObject.prop('hostname'), hrefObject.prop('pathname'));
        }

        /**
         * Return the original url with query string stripped.
         * @param url The URL from which to extract the section.
         * @returns the original url with query string stripped.
         */
        public static getUrlWithoutQueryString(url: string): string {
            let hrefObject = Utility.getHrefObjectFromUrl(url);
            return hrefObject.prop('protocol') + '//' + Utility.urlCombine(hrefObject.prop('host'), hrefObject.prop('pathname'));
        }

        /**
         * Extracts the protocol section of a url.
         * @param url The URL from which to extract the section.
         * @returns The protocol for the current URL.
         */
        public static getProtocolFromUrl(url: string): string {
            return Utility.getHrefObjectFromUrl(url).prop('protocol').replace(':', '');
        }

        /**
         * Returns a formatted href object from a URL.
         * @param url The URL used to generate the object.
         * @returns A jQuery object with the url.
         */
        public static getHrefObjectFromUrl(url: string): JQuery {
            let aObject = $('<a>');
            aObject = aObject.prop('href', url);
            return aObject;
        }

        /**
         * Converts a WCF representation of a dictionary to a JavaScript dictionary.
         * @param wcfDictionary The WCF dictionary to convert.
         * @returns The native JavaScript representation of this dictionary.
         */
        public static convertWcfToJsDictionary(wcfDictionary: any[]): { [index: string]: any; } {
            // convert the WCF JSON representation of a dictionary
            // to JS dictionary.
            // WCF representation: [{"Key": Key, "Value": Value}..]
            // JS representation: [Key: Value ..]

            let result: { [index: string]: any; } = {};

            for (let i = 0; i < wcfDictionary.length; i++) {
                let keyValuePair = wcfDictionary[i];
                result[keyValuePair['Key']] = keyValuePair['Value'];
            }

            return result;
        }

        public static getDateFromWcfJsonString(jsonDate: string, fromUtcMilliseconds: boolean): Date {
            if (StringExtensions.isNullOrEmpty(jsonDate)) {
                return null;
            }
            let begIndex = jsonDate.indexOf('(');
            let endIndex = jsonDate.indexOf(')');
            if (begIndex !== -1 && endIndex !== -1) {
                let milliseconds = parseInt(jsonDate.substring(begIndex + 1, endIndex), 10);

                if (fromUtcMilliseconds) {
                    return new Date(milliseconds);
                }
                else {
                    let retValue = new Date(0);
                    retValue.setUTCMilliseconds(milliseconds);
                    return retValue;
                }
            }
            return null;
        }

        /**
         * Get the outer html of the given jquery object.
         * @param content The jquery object.
         * @returns The entire html representation of the object.
         */
        public static getOuterHtml(content: JQuery): string {
            return $('<div>').append(content).html();
        }

        /**
         * Comparison Method: Compares two integer numbers.
         * @param a An integer value.
         * @param b An integer value.
         * @returns The comparison result.
         */
        public static compareInt(a: number, b: number): number {
            return a - b;
        }
        
        /**
         * Return the index of the smallest value in a numerical array.
         * @param a A numeric array.
         * @returns The index of the smallest value in the array.
         */
        public static getIndexOfMinValue(a: number[]) {
            let retValue = 0;
            let currentMinValue = a[0];

            for (let i = 0; i < a.length; i++) {
                if (a[i] < currentMinValue) {
                    currentMinValue = a[i];
                    retValue = i;
                }
            }

            return retValue;
        }
        
        /**
         * Tests whether a URL is valid.
         * @param url The url to be tested.
         * @returns Whether the provided url is valid.
         */
        public static isValidUrl(url: string): boolean {
            return !StringExtensions.isNullOrEmpty(url) &&
                (StringExtensions.startsWithIgnoreCase(url, 'http://') || StringExtensions.startsWithIgnoreCase(url, 'https://'));
        }
        
        /**
         * Extracts a url from a background image attribute in the format of: url('www.foobar.com/image.png').
         * @param input The value of the background-image attribute.
         * @returns The extracted url.
         */
        public static extractUrlFromCssBackgroundImage(input: string) {
            return input.replace(/"/g, "").replace(/url\(|\)$/ig, "");
        }
        
        /**
         * Verifies image data url of images.
         */
        public static isValidImageDataUrl(url: string): boolean {
            let regex: RegExp = new RegExp('data:(image\/(png|jpg|jpeg|gif|svg))');
            return regex.test(url);
        }
        
        /**
         * Downloads a content string as a file.
         * @param content Content stream.
         * @param fileName File name to use.
         */
        public static saveAsFile(content: any, fileName: string): void {
            let contentBlob = new Blob([content], { type: HttpConstants.ApplicationOctetStream });
            let url = window['webkitURL'] || URL;
            let urlLink = url.createObjectURL(contentBlob);
            let fileNameLink = fileName || urlLink;

            // IE support, use msSaveOrOpenBlob API
            if (window.navigator.msSaveOrOpenBlob) {
                window.navigator.msSaveOrOpenBlob(contentBlob, fileNameLink);
                return;
            }

            // WebKit-based browser support requires generating an anchor tag with
            // download attribute set to blob store and triggering a click event to invoke 
            // a download to file action
            let hyperlink = document.createElement('a');
            hyperlink.href = urlLink;
            hyperlink.target = '_blank';
            hyperlink['download'] = fileNameLink;
            document.body.appendChild(hyperlink);
            hyperlink.click();

            document.body.removeChild(hyperlink);
        }

        /**
         * Helper method to get the simple type name from a typed object.
         * @param obj The typed object.
         * @returns The simple type name for the object.
         */
        public static getType(obj: TypedObject) {
            Utility.throwIfNullEmptyOrWhitespaceString(obj.__type, this, 'getType', 'obj');

            let parts = obj.__type.split(":");

            if (parts.length !== 2) {
                Errors.argument("obj.__type", "Type String not in expected format [Type]#[Namespace]: " + obj.__type);
            }

            if (parts[1] !== Utility.TypeNamespace) {
                Errors.argument("obj.__type", "Type Namespace not expected: " + parts[1]);
            }

            return parts[0];
        }

        /** 
         * Check if an element supports a specific event type.
         * @param eventName The name of the event.
         * @param element The element to test for event support.
         * @returns Whether the even is supported on the provided element.
         */
        public static isEventSupported(eventName: string, element: Element): boolean {
            eventName = 'on' + eventName;
            let isSupported = (eventName in element);

            if (!isSupported) {
                // if we can't use setAttribute try a generic element
                if (!element.setAttribute) {
                    element = document.createElement('div');
                }
                if (element.setAttribute && element.removeAttribute) {
                    element.setAttribute(eventName, '');
                    isSupported = typeof element[eventName] === 'function';

                    // if the property was created - remove it
                    if (typeof element[eventName] !== 'undefined') {
                        element[eventName] = null;
                    }

                    element.removeAttribute(eventName);
                }
            }

            element = null;
            return isSupported;
        }

        public static toPixel(pixelAmount: number): string {
            Utility.throwIfNullOrUndefined(pixelAmount, this, "toPixel", "pixelAmount");

            return pixelAmount.toString() + CssConstants.pixelUnits;
        }

        public static getPropertyCount(object: any) {
            Utility.throwIfNullOrUndefined(object, this, "getPropertyCount", "object");
            return Object.getOwnPropertyNames(object).length;
        }

        /** 
         * Check if an element supports a specific event type.
         * @param filePath File path.
         * @returns File extension.
         */
        public static getFileExtension(filePath: string): string {
            if (filePath) {
                let index = filePath.lastIndexOf('.');
                if (index >= 0)
                    return filePath.substr(index + 1);
            }
            return '';
        }

        /** 
         * Extract the filename out of a full path delimited by '\' or '/'.
         * @param filePath File path.
         * @returns filename File name.
         */
        public static extractFileNameFromPath(filePath: string): string {
            return filePath.replace(/^.*[\\\/]/, '');
        }

        /**
         * This method indicates whether window.clipboardData is supported.
         * For example, clipboard support for Windows Store apps is currently disabled 
         * since window.clipboardData is unsupported (it raises access denied error)
         * since clipboard in Windows Store is being 
         * achieved through Windows.ApplicationModel.DataTransfer.Clipboard class.
         */
        public static canUseClipboard(): boolean {
            return (typeof MSApp === "undefined");
        }

        public static is64BitOperatingSystem(): boolean {
            return navigator.userAgent.indexOf("WOW64") !== -1 ||
                navigator.userAgent.indexOf("Win64") !== -1;
        }

        public static parseNumber(value: any, defaultValue?: number): number {
            if (value === null)
                return null;
            if (value === undefined)
                return defaultValue;
            
            let result = Number(value);
            if (isFinite(result))
                return result;
            if (isNaN(result) && !(typeof value === "number" || value === "NaN"))
                return defaultValue;
            return result;
        }

        public static getURLParamValue(name:string) {
            let results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(window.location.href);
            if (results == null) {
                return null;
            }
            else {
                return results[1] || 0;
            }
        }

        /**
         * Return local timezone.
         * This function uses summer and winter offset to determine local time zone.
         * The result localTimeZoneString must be a subset of the strings used by server, 
         * as documented here: https://msdn.microsoft.com/en-us/library/gg154758.aspx (Dynamic Daylight Savings Time (Compact 2013)).
         * @return Local timezone string or UTC if timezone cannot be found.
         */
        public static getLocalTimeZoneString(): string {
            let timeSummer = new Date(Date.UTC(2005, 6, 30, 0, 0, 0, 0));
            let summerOffset = -1 * timeSummer.getTimezoneOffset();
            let timeWinter = new Date(Date.UTC(2005, 12, 30, 0, 0, 0, 0));
            let winterOffset = -1 * timeWinter.getTimezoneOffset();
            let localTimeZoneString;

            if (-720 === summerOffset && -720 === winterOffset) { localTimeZoneString = 'Dateline Standard Time'; }
            else if (-660 === summerOffset && -660 === winterOffset) { localTimeZoneString = 'UTC-11'; }
            else if (-660 === summerOffset && -660 === winterOffset) { localTimeZoneString = 'Samoa Standard Time'; }
            else if (-600 === summerOffset && -600 === winterOffset) { localTimeZoneString = 'Hawaiian Standard Time'; }
            else if (-480 === summerOffset && -540 === winterOffset) { localTimeZoneString = 'Alaskan Standard Time'; }
            else if (-420 === summerOffset && -480 === winterOffset) { localTimeZoneString = 'Pacific Standard Time'; }
            else if (-420 === summerOffset && -420 === winterOffset) { localTimeZoneString = 'US Mountain Standard Time'; }
            else if (-360 === summerOffset && -420 === winterOffset) { localTimeZoneString = 'Mountain Standard Time'; }
            else if (-360 === summerOffset && -360 === winterOffset) { localTimeZoneString = 'Central America Standard Time'; }
            else if (-300 === summerOffset && -360 === winterOffset) { localTimeZoneString = 'Central Standard Time'; }
            else if (-300 === summerOffset && -300 === winterOffset) { localTimeZoneString = 'SA Pacific Standard Time'; }
            else if (-240 === summerOffset && -300 === winterOffset) { localTimeZoneString = 'Eastern Standard Time'; }
            else if (-270 === summerOffset && -270 === winterOffset) { localTimeZoneString = 'Venezuela Standard Time'; }
            else if (-240 === summerOffset && -240 === winterOffset) { localTimeZoneString = 'SA Western Standard Time'; }
            else if (-240 === summerOffset && -180 === winterOffset) { localTimeZoneString = 'Central Brazilian Standard Time'; }
            else if (-180 === summerOffset && -240 === winterOffset) { localTimeZoneString = 'Atlantic Standard Time'; }
            else if (-180 === summerOffset && -180 === winterOffset) { localTimeZoneString = 'Montevideo Standard Time'; }
            else if (-180 === summerOffset && -120 === winterOffset) { localTimeZoneString = 'E. South America Standard Time'; }
            else if (-150 === summerOffset && -210 === winterOffset) { localTimeZoneString = 'Mid-Atlantic Standard Time'; }
            else if (-120 === summerOffset && -120 === winterOffset) { localTimeZoneString = 'SA Eastern Standard Time'; }
            else if (0 === summerOffset && 0 === winterOffset) { localTimeZoneString = 'UTC'; }
            else if (60 === summerOffset && 0 === winterOffset) { localTimeZoneString = 'GMT Standard Time'; }
            else if (60 === summerOffset && 120 === winterOffset) { localTimeZoneString = 'Namibia Standard Time'; }
            else if (120 === summerOffset && 60 === winterOffset) { localTimeZoneString = 'Romance Standard Time'; }
            else if (120 === summerOffset && 120 === winterOffset) { localTimeZoneString = 'South Africa Standard Time'; }
            else if (180 === summerOffset && 120 === winterOffset) { localTimeZoneString = 'GTB Standard Time'; }
            else if (180 === summerOffset && 180 === winterOffset) { localTimeZoneString = 'E. Africa Standard Time'; }
            else if (240 === summerOffset && 180 === winterOffset) { localTimeZoneString = 'Russian Standard Time'; }
            else if (240 === summerOffset && 240 === winterOffset) { localTimeZoneString = 'Arabian Standard Time'; }
            else if (270 === summerOffset && 210 === winterOffset) { localTimeZoneString = 'Iran Standard Time'; }
            else if (270 === summerOffset && 270 === winterOffset) { localTimeZoneString = 'Afghanistan Standard Time'; }
            else if (300 === summerOffset && 240 === winterOffset) { localTimeZoneString = 'Pakistan Standard Time'; }
            else if (300 === summerOffset && 300 === winterOffset) { localTimeZoneString = 'West Asia Standard Time'; }
            else if (330 === summerOffset && 330 === winterOffset) { localTimeZoneString = 'India Standard Time'; }
            else if (345 === summerOffset && 345 === winterOffset) { localTimeZoneString = 'Nepal Standard Time'; }
            else if (360 === summerOffset && 300 === winterOffset) { localTimeZoneString = 'N. Central Asia Standard Time'; }
            else if (360 === summerOffset && 360 === winterOffset) { localTimeZoneString = 'Central Asia Standard Time'; }
            else if (390 === summerOffset && 390 === winterOffset) { localTimeZoneString = 'Myanmar Standard Time'; }
            else if (420 === summerOffset && 360 === winterOffset) { localTimeZoneString = 'North Asia Standard Time'; }
            else if (420 === summerOffset && 420 === winterOffset) { localTimeZoneString = 'SE Asia Standard Time'; }
            else if (480 === summerOffset && 420 === winterOffset) { localTimeZoneString = 'North Asia East Standard Time'; }
            else if (480 === summerOffset && 480 === winterOffset) { localTimeZoneString = 'China Standard Time'; }
            else if (540 === summerOffset && 480 === winterOffset) { localTimeZoneString = 'Yakutsk Standard Time'; }
            else if (540 === summerOffset && 540 === winterOffset) { localTimeZoneString = 'Tokyo Standard Time'; }
            else if (570 === summerOffset && 570 === winterOffset) { localTimeZoneString = 'Cen. Australia Standard Time'; }
            else if (600 === summerOffset && 600 === winterOffset) { localTimeZoneString = 'E. Australia Standard Time'; }
            else if (600 === summerOffset && 660 === winterOffset) { localTimeZoneString = 'AUS Eastern Standard Time'; }
            else if (660 === summerOffset && 600 === winterOffset) { localTimeZoneString = 'Tasmania Standard Time'; }
            else if (660 === summerOffset && 660 === winterOffset) { localTimeZoneString = 'West Pacific Standard Time'; }
            else if (690 === summerOffset && 690 === winterOffset) { localTimeZoneString = 'Central Pacific Standard Time'; }
            else if (720 === summerOffset && 660 === winterOffset) { localTimeZoneString = 'Magadan Standard Time'; }
            else if (720 === summerOffset && 720 === winterOffset) { localTimeZoneString = 'Fiji Standard Time'; }
            else if (720 === summerOffset && 780 === winterOffset) { localTimeZoneString = 'New Zealand Standard Time'; }
            else if (780 === summerOffset && 780 === winterOffset) { localTimeZoneString = 'Tonga Standard Time'; }
            else { localTimeZoneString = 'UTC'; }
            return localTimeZoneString;
        }
    }

    export class VersionUtility {
        /**
         * Compares 2 version strings.
         * @param versionA The first version string.
         * @param versionB The second version string.
         * @returns A result for the comparison.
         */
        static compareVersions(versionA: string, versionB: string): number {
            let a = versionA.split('.').map(parseFloat);
            let b = versionB.split('.').map(parseFloat);

            let versionParts = Math.max(a.length, b.length);

            for (let i = 0; i < versionParts; i++) {
                let partA = a[i] || 0;
                let partB = b[i] || 0;

                if (partA > partB)
                    return 1;

                if (partA < partB)
                    return -1;
            }

            return 0;
        }
    }

    export module PerformanceUtil {
        export class PerfMarker {
            private _name: string;
            private _start: string;

            constructor(name: string) {
                this._name = name;
                this._start = PerfMarker.begin(name);
            }

            private static begin(name: string) {
                if (window.performance === undefined || performance.mark === undefined) return;
                if (console.time) {
                    console.time(name);
                }
                name = 'Begin ' + name;
                performance.mark(name);
                return name;
            }

            public end() {
                if (window.performance === undefined || performance.mark === undefined || performance.measure === undefined) return;
                let name = this._name;
                let end = 'End ' + name;
                performance.mark(end);
                performance.measure(name, this._start, end);
                if (console.timeEnd) {
                    console.timeEnd(name);
                }
            }
        }

        export function create(name: string): PerfMarker {
            return new PerfMarker(name);
        }
    }

    export module DeferUtility {        
        /**
         * Wraps a callback and returns a new function.
         * The function can be called many times but the callback
         * will only be executed once on the next frame.
         * Use this to throttle big UI updates and access to DOM.
         */
        export function deferUntilNextFrame(callback: Function): Function {
            let isWaiting, args, context;

            if (!window.requestAnimationFrame) {
                window.requestAnimationFrame = (func) => setTimeout(func, 1000 / 50);
            }

            return function() {
                if (!isWaiting) {
                    isWaiting = true;
                    args = arguments;
                    context = this;
                    window.requestAnimationFrame(() => {
                        isWaiting = false;
                        callback.apply(context, args);
                    });
                }
            };
        }
    }
}
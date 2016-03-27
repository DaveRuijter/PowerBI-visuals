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

module jsCommon {
    // NOTE: this file includes standalone utilities that should have no dependencies on external libraries, including jQuery.

    /**
     * Interface to help define objects indexed by number to a particular type.
     */
    export interface INumberDictionary<T> {
        [key: number]: T;
    }

    /**
     * Interface to help define objects indexed by name to a particular type.
     */
    export interface IStringDictionary<T> {
        [key: string]: T;
    }

    /**
     * Extensions for Enumerations.
     */
    export module EnumExtensions {
        /**
         * Gets a value indicating whether the value has the bit flags set.
         */
        export function hasFlag(value: number, flag: number): boolean {
            debug.assert(!!flag, 'flag must be specified and nonzero.');

            return (value & flag) === flag;
        }

        /**
         * Sets a value of a flag without modifying any other flags.
         */        
        export function setFlag(value: number, flag: number): number {
            debug.assert(!!flag, "flag must be specified and nonzero.");
            return value |= flag;
        }

        /**
         * Resets a value of a flag without modifying any other flags.
         */                
        export function resetFlag(value: number, flag: number): number {
            debug.assert(!!flag, "flag must be specified and nonzero.");
            return value &= ~flag;
        }

        /**
         * According to the TypeScript Handbook, this is safe to do.
         */
        export function toString(enumType: any, value: number): string {
            return enumType[value];
        }
    }

    /**
     * Extensions to String class.
     */
    export module StringExtensions {
        /**
         * Checks if a string ends with a sub-string.
         */
        export function endsWith(str: string, suffix: string): boolean {
            debug.assertValue(str, 'str');
            debug.assertValue(suffix, 'suffix');

            return str.indexOf(suffix, str.length - suffix.length) !== -1;
        }
    }

    export module LogicExtensions {
        export function XOR(a: boolean, b: boolean): boolean {
            return (a || b) && !(a && b);
        }
    }

    export module JsonComparer {
        /**
         * Performs JSON-style comparison of two objects.
         */
        export function equals<T>(x: T, y: T): boolean {
            if (x === y)
                return true;

            return JSON.stringify(x) === JSON.stringify(y);
        }
    }

    /**
     * Values are in terms of 'pt'
     * Convert to pixels using PixelConverter.fromPoint
     */
    export module TextSizeDefaults {
        /**
         * Stored in terms of 'pt'
         * Convert to pixels using PixelConverter.fromPoint
         */
        export const TextSizeMin: number = 8;

        /**
         * Stored in terms of 'pt'
         * Convert to pixels using PixelConverter.fromPoint
         */
        export const TextSizeMax: number = 40;

        const TextSizeRange: number = TextSizeMax - TextSizeMin;

        /**
         * Returns the percentage of this value relative to the TextSizeMax
         * @param textSize - should be given in terms of 'pt'
         */
        export function getScale(textSize: number) {
            return (textSize - TextSizeMin) / TextSizeRange;
        }

    }

    export module PixelConverter {
        const PxPtRatio: number = 4 / 3;
        const PixelString: string = 'px';

        /**
         * Appends 'px' to the end of number value for use as pixel string in styles
         */
        export function toString(px: number): string {
            return px + PixelString;
        }

        /**
         * Converts point value (pt) to pixels
         */
        export function fromPoint(pt: number): string {
            return toString(PxPtRatio * pt);
        }
    }
}

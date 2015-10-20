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

module powerbi.visuals {
    import IStringResourceProvider = jsCommon.IStringResourceProvider;

    export class NoMapLocationWarning implements IVisualWarning {
        public get code(): string {
            return 'NoMapLocation';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            let messageKey: string = 'NoMapLocationMessage';
            let titleKey: string = 'NoMapLocationKey';
            let detailKey: string = 'NoMapLocationValue';

            let visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: resourceProvider.get(titleKey),
                detail: resourceProvider.get(detailKey),
            };

            return visualMessage;
        }
    }

    export class FilledMapWithoutValidGeotagCategoryWarning implements IVisualWarning {
        public get code(): string {
            return 'NoValidGeotaggedCategory';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            let messageKey: string = 'NoValidGeotaggedCategoryMessage';
            let titleKey: string = 'NoValidGeotaggedCategoryKey';
            let detailKey: string = 'NoValidGeotaggedCategoryValue';

            let visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: resourceProvider.get(titleKey),
                detail: resourceProvider.get(detailKey),
            };

            return visualMessage;
        }
    }

    export class GeometryCulledWarning implements IVisualWarning {
        public get code(): string {
            return 'GeometryCulledWarning';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            let messageKey: string = 'DsrLimitWarning_TooMuchDataMessage';
            let titleKey: string = '';
            let detailKey: string = '';

            let visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: resourceProvider.get(titleKey),
                detail: resourceProvider.get(detailKey),
            };

            return visualMessage;
        }
    }

    export class NegativeValuesNotSupportedWarning implements IVisualWarning {
        public get code(): string {
            return 'NegativeValuesNotSupported';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            var messageKey: string = 'VisualWarning_NegativeValues';

            var visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: '',
                detail: '',
            };

            return visualMessage;
        }
    }

    export class AllNegativeValuesWarning implements IVisualWarning {
        public get code(): string {
            return 'AllNegativeValuesNotSupported';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            var messageKey: string = 'VisualWarning_AllNegativeValues';

            var visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: '',
                detail: '',
            };

            return visualMessage;
        }
    }

    export class NaNNotSupportedWarning implements IVisualWarning {
        public get code(): string {
            return 'NaNNotSupported';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            let messageKey: string = 'VisualWarning_InfinityValues';

            let visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: '',
                detail: '',
            };

            return visualMessage;
        }
    }

    export class InfinityValuesNotSupportedWarning implements IVisualWarning {
        public get code(): string {
            return 'InfinityValuesNotSupported';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            let messageKey: string = 'VisualWarning_InfinityValues';

            let visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: '',
                detail: '',
            };

            return visualMessage;
        }
    }

    export class ValuesOutOfRangeWarning implements IVisualWarning {
        public get code(): string {
            return 'ValuesOutOfRange';
        }

        public getMessages(resourceProvider: IStringResourceProvider): IVisualErrorMessage {
            let messageKey: string = 'VisualWarning_VisualizationOutOfRange';

            let visualMessage: IVisualErrorMessage = {
                message: resourceProvider.get(messageKey),
                title: '',
                detail: '',
            };

            return visualMessage;
        }
    }
}
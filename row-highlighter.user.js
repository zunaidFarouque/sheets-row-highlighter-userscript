// ==UserScript==
// @name         Sheets Row Highlighter (Lightweight)
// @namespace    https://github.com/zunaidFarouque/sheets-row-highlighter-userscript
// @version      1.0.0
// @description  Highlight row/column in Google Sheets & Excel Online. Lightweight port of the original extension.
// @author       Md Zunaid Farouque (Original extension by matsu7089)
// @license      MIT
// @match        https://docs.google.com/spreadsheets/d/*
// @match        https://*.officeapps.live.com/x/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // ==========================================
    // USER CONFIGURATION (Change these values)
    // ==========================================
    const CONFIG = {
        color: '#0e65eb', // Hex color code
        opacity: '0.1',   // 0.01 to 0.5 recommended
        enableRow: true,  // true or false
        enableCol: false  // true or false
    };
    // ==========================================

    // --- Core Logic (From Original Extension) ---

    class SheetsActiveCellLocator {
        _activeBorderClass = 'active-cell-border';
        _selectionClass = 'selection';
        _sheetContainerId = 'waffle-grid-container';

        getHighlightRectList() {
            const activeSelectionList = Array.from(
                document.getElementsByClassName(this._selectionClass)
            ).filter((element) => element.style.display !== 'none');

            if (activeSelectionList.length) {
                return this._getMultipleHighlightRectList(activeSelectionList);
            }
            return this._getSingleHighlightRectList();
        }

        _getMultipleHighlightRectList(activeSelectionList) {
            const sheetRect = this._getSheetContainerRect();
            if (!sheetRect) return [];

            const activeSelectionRectList = activeSelectionList.map((element) => {
                const { x, y, width, height } = element.getBoundingClientRect();
                return {
                    x: Math.ceil(x - sheetRect.x),
                    y: Math.ceil(y - sheetRect.y),
                    width: Math.ceil(width),
                    height: Math.ceil(height),
                };
            });

            const rowOrColumnRectList = activeSelectionRectList.filter(
                (rect) => sheetRect.width < rect.width || sheetRect.height < rect.height
            );

            return activeSelectionRectList.filter(
                (rect) =>
                    !rowOrColumnRectList.some(({ x, y, width, height }) =>
                        height < width
                            ? rect.y === y && rect.height === height
                            : rect.x === x && rect.width === width
                    )
            );
        }

        _getSingleHighlightRectList() {
            const sheetRect = this._getSheetContainerRect();
            const activeBorderList = document.getElementsByClassName(this._activeBorderClass);

            if (!sheetRect || activeBorderList.length !== 4) return [];

            const topBorderRect = activeBorderList[0].getBoundingClientRect();
            const leftBorderRect = activeBorderList[3].getBoundingClientRect();

            return [{
                x: topBorderRect.x - sheetRect.x,
                y: topBorderRect.y - sheetRect.y,
                width: topBorderRect.width,
                height: leftBorderRect.height,
            }];
        }

        _getSheetContainerRect() {
            return document.getElementById(this._sheetContainerId)?.getBoundingClientRect();
        }

        getSheetContainerStyle() {
            const { x, y, width, height } = this._getSheetContainerRect() || {};
            return {
                left: `${x}px`,
                top: `${y}px`,
                width: `${width}px`,
                height: `${height}px`,
            };
        }
    }

    class ExcelActiveCellLocator {
        _singleSelectionClassList = ['ewr-cell-selection-highlight-all-after-fluent', 'ewr-cell-selection-highlight-all'];
        _multipleSelectionClassList = ['ewr-discontinuous-selection-active-range-border-after-fluent', 'ewr-discontinuous-selection-active-range-border', 'ewr-discontinuous-selection-after-fluent', 'ewr-discontinuous-selection'];
        _sheetContainerClass = 'ewa-grid-ltr';
        _hiddenClass = 'ewa-hidden';

        getHighlightRectList() {
            const sheetRect = this._getSheetContainerRect();
            if (!sheetRect) return [];

            const singleSelectionList = this._singleSelectionClassList
                .flatMap((className) => Array.from(document.getElementsByClassName(className)))
                .filter((element) => !element.classList.contains(this._hiddenClass));

            const rectList = singleSelectionList.length
                ? this._getSingleHighlightRectList(singleSelectionList)
                : this._getMultipleHighlightRectList();

            return rectList.filter(
                (rect) => !(sheetRect.width < rect.width || sheetRect.height < rect.height)
            );
        }

        _getSingleHighlightRectList(singleSelectionList) {
            const sheetRect = this._getSheetContainerRect();
            if (!sheetRect) return [];

            const selectionRect = singleSelectionList.sort((a, b) => a.id.localeCompare(b.id))[0].getBoundingClientRect();

            return [{
                x: selectionRect.x - sheetRect.x,
                y: selectionRect.y - sheetRect.y,
                width: selectionRect.width,
                height: selectionRect.height,
            }];
        }

        _getMultipleHighlightRectList() {
            const sheetRect = this._getSheetContainerRect();
            if (!sheetRect) return [];

            const selectionList = this._multipleSelectionClassList
                .flatMap((className) => Array.from(document.getElementsByClassName(className)))
                .filter((element) => !element.classList.contains(this._hiddenClass));

            return selectionList.map((element) => {
                const { x, y, width, height } = element.getBoundingClientRect();
                return { x: x - sheetRect.x, y: y - sheetRect.y, width, height };
            });
        }

        _getSheetContainerRect() {
            return document.getElementsByClassName(this._sheetContainerClass)[0]?.getBoundingClientRect();
        }

        getSheetContainerStyle() {
            const sheetContainer = document.getElementsByClassName(this._sheetContainerClass)[0];
            const { x, y, width, height } = sheetContainer?.getBoundingClientRect() || {};
            return {
                left: `${x}px`,
                top: `${y}px`,
                width: `${width}px`,
                height: `${height}px`,
                zoom: sheetContainer?.style['zoom'],
                zIndex: '1',
            };
        }
    }

    class RowHighlighterApp {
        constructor(appContainer, locator) {
            this.appContainer = appContainer;
            this.locator = locator;
            this.elementPool = [];
            
            // Load from user config block
            this.backgroundColor = CONFIG.color;
            this.opacity = CONFIG.opacity;
            this.isRowEnabled = CONFIG.enableRow;
            this.isColEnabled = CONFIG.enableCol;
        }

        update() {
            const rectList = this.locator.getHighlightRectList();

            Object.assign(this.appContainer.style, {
                position: 'absolute',
                pointerEvents: 'none',
                overflow: 'hidden',
                ...this.locator.getSheetContainerStyle(),
            });

            const highlightTaskList = (
                this.isRowEnabled
                    ? this._mergeRectList(rectList, 'y').map(({ height, y }) => ({
                        left: '0px', top: `${y}px`, width: '100%', height: `${height}px`,
                    })) : []
            ).concat(
                this.isColEnabled
                    ? this._mergeRectList(rectList, 'x').map(({ width, x }) => ({
                        left: `${x}px`, top: '0px', width: `${width}px`, height: '100%',
                    })) : []
            );

            const diff = highlightTaskList.length - this.elementPool.length;

            if (0 < diff) {
                Array.from({ length: diff }).forEach(() => {
                    const element = document.createElement('div');
                    this.elementPool.push(element);
                    this.appContainer.appendChild(element);
                });
            }

            if (diff < 0) {
                this.elementPool.slice(diff).forEach((element) => { element.style.display = 'none'; });
            }

            highlightTaskList.forEach((task, index) => {
                const element = this.elementPool[index];
                Object.assign(element.style, {
                    position: 'absolute', pointerEvents: 'none', display: 'block',
                    backgroundColor: this.backgroundColor, opacity: this.opacity,
                    ...task,
                });
            });
        }

        _mergeRectList(rectList, dim) {
            return [...rectList]
                .sort((a, b) => a[dim] - b[dim])
                .reduce((acc, rect) => {
                    const prevRect = acc[acc.length - 1];
                    const dimSize = dim === 'x' ? 'width' : 'height';

                    if (!prevRect || prevRect[dim] + prevRect[dimSize] < rect[dim]) {
                        acc.push({ ...rect });
                        return acc;
                    }

                    prevRect[dimSize] = Math.max(prevRect[dimSize], rect[dim] + rect[dimSize] - prevRect[dim]);
                    return acc;
                }, []);
        }
    }

    // --- Initialization ---
    const appContainer = document.createElement('div');
    appContainer.id = 'rh-app-container';
    document.body.appendChild(appContainer);

    const locator = location.host === 'docs.google.com' 
        ? new SheetsActiveCellLocator() 
        : new ExcelActiveCellLocator();

    const app = new RowHighlighterApp(appContainer, locator);
    const updateHighlight = app.update.bind(app);

    // Event Listeners
    window.addEventListener('click', updateHighlight);
    window.addEventListener('keydown', updateHighlight);
    window.addEventListener('keyup', updateHighlight);
    window.addEventListener('resize', updateHighlight);
    window.addEventListener('scroll', updateHighlight, true);

    // Initial draw
    updateHighlight();

})();

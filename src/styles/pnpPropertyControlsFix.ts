/**
 * Fix for @pnp/spfx-property-controls PropertyFieldCollectionData table layout.
 *
 * SPFx 1.22's sp-css-loader applies postcss-modules-scope to .module.css files,
 * re-hashing class names. The PnP package ships pre-compiled CSS with baked-in
 * class names (e.g. table_f8375039) and hardcoded JS mappings. The re-hashing
 * causes a mismatch: the DOM uses the original class names but the injected CSS
 * has different (re-hashed) selectors, so display:table/table-row/table-cell
 * never applies and the collection data panel renders fields vertically.
 *
 * This module injects the original CSS rules directly via a <style> tag,
 * guaranteeing the styles match the class names in the DOM.
 */

const STYLE_TAG_ID = 'sp-search-pnp-property-controls-fix';

let injected = false;

export function ensurePnpPropertyControlStyles(): void {
  if (injected) {
    return;
  }
  if (typeof document === 'undefined') {
    return;
  }
  if (document.getElementById(STYLE_TAG_ID)) {
    injected = true;
    return;
  }

  const style = document.createElement('style');
  style.id = STYLE_TAG_ID;
  style.textContent = PNP_COLLECTION_DATA_CSS;
  document.head.appendChild(style);
  injected = true;
}

/**
 * Critical CSS from @pnp/spfx-property-controls PropertyFieldCollectionDataHost.module.css.
 * These rules use the baked-in hash suffix _f8375039 that matches the JS class name mapping.
 */
const PNP_COLLECTION_DATA_CSS = `
.collectionData_f8375039 {
  background-color: #edebe9;
  padding: 10px;
}
.noCollectionData_f8375039 {
  text-align: center;
}
.panelActions_f8375039 {
  margin-top: 15px;
  text-align: right;
}
.panelActions_f8375039 button {
  margin-right: 15px;
}
.panelActions_f8375039 button:last-child {
  margin-right: 0;
}
.required_f8375039 {
  color: #d83b01;
  font-size: 8px;
  vertical-align: super;
}
.addBtn_f8375039 {
  color: #0078d4;
}
.addBtnDisabled_f8375039 {
  color: #d2d0ce;
}
.inputField_f8375039 {
  color: inherit;
}
.numberField_f8375039 {
  box-sizing: border-box;
  box-shadow: none;
  margin: 0;
  padding: 0;
  border: 1px solid #a19f9d;
  background: #fff;
  height: 32px;
  display: flex;
  flex-direction: row;
  align-items: stretch;
  position: relative;
}
.numberField_f8375039:hover {
  border-color: #201f1e;
}
.numberField_f8375039 input {
  box-sizing: border-box;
  box-shadow: none;
  margin: 0;
  font-size: 14px;
  border-radius: 0;
  border: none;
  background: transparent;
  color: #323130;
  padding: 0 12px;
  width: 100%;
  text-overflow: ellipsis;
  outline: 0;
}
.numberField_f8375039.invalidField_f8375039 {
  border-color: #a80000;
}
.collectionDataField_f8375039 > span {
  display: none;
}
.collectionDataField_f8375039 div[class^=invalid_] {
  border-color: #a80000 !important;
}
.table_f8375039 {
  display: table;
  width: 100%;
  border-collapse: collapse;
}
.tableRow_f8375039 {
  display: table-row;
  line-height: 30px;
}
.tableRow_f8375039:hover {
  background-color: #f3f2f1;
  cursor: pointer;
  outline: 1px solid transparent;
}
.tableRow_f8375039.tableFooter_f8375039 {
  background-color: #edebe9;
  border-top: 1px solid #a19f9d;
}
.tableCell_f8375039 {
  display: table-cell;
  padding: 0 10px;
  vertical-align: middle;
}
.tableCell_f8375039 > div {
  margin-bottom: 8px;
  margin-top: 8px;
}
.tableCell_f8375039 > div.ms-TextField {
  margin-bottom: 8px;
  margin-top: 8px;
}
.errorCallout_f8375039 {
  padding: 0 15px;
  min-width: 200px;
}
.errorCalloutLink_f8375039:not([disabled]) {
  color: #a80000;
}
.errorMsgs_f8375039 {
  font-size: 14px;
  font-weight: 400;
}
.errorMsgs_f8375039 p {
  font-size: 17px;
  font-weight: 300;
}
.errorMsgs_f8375039 ul {
  padding-left: 15px;
}
.errorMsgs_f8375039 li {
  color: #a80000;
}
.tableHead_f8375039 {
  font-weight: 300;
  font-size: 12px;
  color: #605e5c;
}
.tableHead_f8375039 .tableCell_f8375039 {
  font-weight: 400;
  text-align: left;
  border-bottom: 1px solid #a19f9d;
}
`;

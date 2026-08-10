import './browser-polyfill';
export * from './index';
// `export *` skips the default export, so re-export it explicitly to keep this
// entry point interchangeable with the main one.
export { default } from './index';

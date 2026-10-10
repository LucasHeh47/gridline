import './styles.css';
import { App } from './app';

const root = document.getElementById('app')!;
const app = new App(root);
// Handy for debugging from DevTools
(window as unknown as { gridline_app: App }).gridline_app = app;

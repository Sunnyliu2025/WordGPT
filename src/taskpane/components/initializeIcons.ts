import { initializeIcons } from "@fluentui/font-icons-mdl2";

/**
 * 初始化 Fluent UI 图标
 *
 * 注意：initializeIcons() 内部通过 Stylesheet.insertRule() 动态注入 @font-face 规则。
 * 在 Office Add-in WebView 中，insertRule() 对 @font-face 的支持可能失败并抛出异常。
 * 因此需要用 try-catch 包裹，防止模块加载级崩溃导致整个 UI 不渲染。
 *
 * CSS 中已提供静态 @font-face 声明作为后备（见 taskpane.css），
 * 即使此函数失败，图标字体仍可通过 CSS 加载。
 */
/* global console */
try {
  initializeIcons();
} catch (e) {
  console.warn("Fluent UI icon initialization failed (expected in some WebView environments):", e);
}

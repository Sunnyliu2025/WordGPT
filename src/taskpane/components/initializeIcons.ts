import { initializeIcons } from "@fluentui/font-icons-mdl2";

// 确保图标只初始化一次
if (!window.iconsInitialized) {
  initializeIcons();
  window.iconsInitialized = true;
}

// 声明全局变量类型
declare global {
  interface Window {
    iconsInitialized: boolean;
  }
} 
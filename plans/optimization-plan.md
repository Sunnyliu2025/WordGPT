# WordGPT 代码优化计划（修正版）

> 根据反馈：DeepSeek 模型已更新，`deepseek-v4-flash` 为有效模型名，不做修改。

---

## 阶段 1: 修复阻塞性 Bug

| # | 文件 | 修改内容 | 风险 |
|---|------|----------|------|
| 1.1 | [`package.json`](package.json) | `dependencies` 中添加 `axios`（代码已`import`但未声明） | 低 |
| 1.2 | [`package.json`](package.json) | `devDependencies` 中添加 `react-hot-loader`（webpack 使用中但缺失） | 低 |
| 1.3 | [`package.json`](package.json) | 移除未使用的 `react-icons` | 低 |
| 1.4 | [`package.json`](package.json) | 添加 `@hot-loader/react-dom` 提升 HMR 性能 | 低 |
| 1.5 | [`package.json`](package.json) | `browserslist` 补充 `last 2 versions` | 低 |

## 阶段 2: 修复 App.tsx 代码健壮性

| # | 文件 | 修改内容 | 风险 |
|---|------|----------|------|
| 2.1 | [`App.tsx:3`](/src/taskpane/components/App.tsx:3) | `import axios` → `import axios, { AxiosError }` | 低 |
| 2.2 | [`App.tsx:73`](/src/taskpane/components/App.tsx:73) | 添加 optional chaining: `response.data?.choices?.[0]?.message?.content` | 低 |
| 2.3 | [`App.tsx:78`](/src/taskpane/components/App.tsx:78) | `catch (error: any)` → `catch (err: unknown)` 并使用 `AxiosError` 类型守卫 | 低 |
| 2.4 | [`App.tsx:95-101`](/src/taskpane/components/App.tsx:95) | `onInsert` 添加 try/catch 错误处理 | 低 |
| 2.5 | [`App.tsx:103-105`](/src/taskpane/components/App.tsx:103) | `onCopy` 添加 await + try/catch 错误处理 | 低 |
| 2.6 | [`App.tsx`](/src/taskpane/components/App.tsx) | 提取常量（API URL、最大长度等），避免魔法值 | 低 |

## 阶段 3: 清理冗余代码

| # | 文件 | 修改内容 | 风险 |
|---|------|----------|------|
| 3.1 | [`global.d.ts`](/src/taskpane/components/global.d.ts) | 删除空文件（`dom` lib 已在 tsconfig 中声明） | 低 |
| 3.2 | [`initializeIcons.ts`](/src/taskpane/components/initializeIcons.ts) | 移除全局 `window.iconsInitialized` 污染（Fluent UI 内部已做去重） | 低 |
| 3.3 | [`index.html:13-14`](index.html:13-14) | 修复 href 链接换行断裂问题 | 低 |

## 阶段 4: 构建配置优化

| # | 文件 | 修改内容 | 风险 |
|---|------|----------|------|
| 4.1 | [`tsconfig.json`](tsconfig.json) | `target: "es5"` → `"es2015"`（Office WebView 支持 ES6+） | 低 |
| 4.2 | [`tsconfig.json`](tsconfig.json) | 添加 `noUnusedLocals: true` | 低 |

## 总结

- **不变更的内容**: DeepSeek 模型名 `deepseek-v4-flash`（官方已更新）
- **不变更的内容**: `Center.tsx` / `Container.tsx` 布局组件（保留清晰抽象）
- **高优先级**: 阶段 1（依赖修复）+ 阶段 2（代码健壮性）
- **低风险**: 所有修改均为单向改进，不改变现有功能逻辑

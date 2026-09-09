# VML 去配置化同步与逻辑审核

审核日期：2026-09-09。工程：`D:\VML\VmlProject-github\VML2HTMLTools`。

## 范围与结论

以 `D:\VML\vmlparse.js` 为基准，并保留用户后续明确的两项要求：先 replaceVMLSrcs 再 changeVMLSrcs；table 只有在最近 VML 容器内部才参与处理，图片最多提升到 table 同级。

本次已删除 DEFAULT_OPTIONS、normalizeOptions、配置驱动的处理分支与 VmlParserOptions，更新 SDK、VmlHost、entry 和使用文档。运行时与 `D:\VML\vmlparse.refactored.js` 字节一致，已在本工程根目录生成同名 JS；原始 vmlparse.js 未修改。

**去配置化已完成，但现有重构实现不能认定为与原始代码完全等价。** 以下差异为本次审核发现，未混入同步修改。修复建议均以恢复原始处理语义、保留已确认的 table 限制为前提，不建议增加策略选项。

当前公开调用：

```typescript
await sdk.parseVMLImagesFromHtml(html);
await host.parseFile(inputPath, outputPath);
```

目标工程此前已收敛为字符串 SDK + VmlHost 文件入口，本次保留这一接口结构，没有重新增加 selector 入口。

## 已确认的问题

### 1. P1：源路径处理超出原始 split('/') 行为

位置：`outlook-vml-har/tools/runtime/vml-runtime.iife.js:262`，`getResourceFileName`。

原始 changeVMLSrcs 只取 `src.split('/').pop()`；当前函数额外移除 cid 前缀、查询参数与片段、解码百分号、转换反斜杠，还会拒绝部分协议和超过 2048 字符的源。这会改变最终写入 VML 的 src。

| 输入 | 原始处理结果 | 当前结果 |
| --- | --- | --- |
| `cid:image001.jpg@id` | `cid:image001.jpg@id` | `image001.jpg@id` |
| `/mail/1/inline/a.png?token=1` | `a.png?token=1` | `a.png` |
| `/mail/1/inline/a%20b.png` | `a%20b.png` | `a b.png` |

以上差异已通过执行当前辅助函数复现。建议恢复原始取文件名操作；空值等边界可以报告错误，但不应悄悄改写合法原值。这也符合用户此前“cid 不用处理”的要求。

### 2. P1：全局 shape 去重改变匹配与聚合结果

位置：运行时 `buildShapeIndex:435`、`createImagePlan:518`。

原始 parseVMLImages 在每个候选容器内收集 shape，再遍历全部 img。同一容器存在两个同 ID shape 时，原始 filter 保留两者，并对各字段取最小值。当前全局 Map 对同 ID 只保留第一条。已复现同 ID 的 left 为 100pt、10pt 时，当前只剩 100pt，原始聚合为 10pt。

跨容器重复 ID 也会使后一个容器的信息被丢弃。不能简单改为“只匹配图片最近容器”，因为原始代码遍历的是 `$content.find('img')`。建议恢复逐容器的处理顺序和 shape 列表语义；重复 ID 可以报告，但不能以报告为由丢弃记录。

### 3. P1：changeVMLSrcs 的匹配数量与覆盖顺序不同

位置：运行时 `rewriteVmlImageSources:784`，特别是 `imageSourceByShapeId.has:796`。

原始代码使用 `vShapesAttr.includes(shape.id)`，只更新 `matchedShapes[0]`；多张 img 对应同一 shape 时，后面的 img 覆盖前面的 src。当前按空格拆分精确 ID，并给每一个 ID 建立映射，同时保留第一个源。

已复现同一 ID 的两张图片依次为 first.png、last.png：原始最终使用 last.png，当前使用 first.png。多 ID、ID 子串匹配也存在差异。建议按原始匹配和赋值顺序处理；严格等价要求下，不应自行改成精确匹配或首次优先。

### 4. P1：li 的偏移特例被应用到所有包装策略

位置：运行时 `computeLayout:475`；对照原始代码约 273、311、352 行。

原始代码仅在已有 span 的处理分支对 li 使用未扣减的坐标；新建 img 包装 span、table 内链接包装 span 的分支仍扣减容器 margin/text-indent。当前 computeLayout 在选择策略之前统一计算，所有 li 都不扣减。

已复现容器 margin-left=20pt、shape left=100pt：当前计算为 100pt；原始新建包装分支应为 80pt。建议先确定原始分支，再决定是否应用 li 特例，不改变已有 span 的清空样式规则。

### 5. P2：无 table 的新建包装节点没有按原逻辑归位

位置：运行时 `applyPlan:667`，节点移动段约 706–714 行；原始代码约 372–379 行。

例如 `VML div > section > img`：当前包裹 img 后，只要节点仍在 VML 容器内，就留在 section 下。原始新包装分支会移到容器第一个直接 span 内，找不到时追加到容器本身。若 section 建立了定位上下文，图片位置也会变化。

建议恢复无 table 分支的原始归位行为；存在 table 时继续遵守已明确要求的“只到 table 同级”和容器边界，不能把此次建议扩展为把所有图片提到 div 下。

### 6. P1：字符集声明替换破坏引号，且可能误改正文或注释

位置：`outlook-vml-har/src/main/ets/VmlText.ets:105`。

normalizeCharsetDeclaration 对全文使用正则替换 charset，匹配掉开始引号却不匹配结束引号。已复现：`<meta charset="gb2312">` 输出 `<meta charset=utf-8">`。此外正文、脚本、注释中的同样文字也可能被替换，超出“仅调整编码声明”的范围。

建议只处理真实 meta 元素的 charset / http-equiv Content-Type 声明，并保留属性引号、正文和 VML 注释。这属于文件适配层修复，不改变 vmlparse.js 的图片处理逻辑。

## 其他边界与待确认项

- **P2，运行时版本缓存**：内核内容和函数签名已改变，但 VERSION 仍是 1.0.0；SDK `ensureRuntime`（OutlookVmlSdk.ets:57）只比较版本。同一 ArkWeb 若已有旧版 1.0.0，会跳过注入，导致新 SDK 调到旧签名。新建页面不受此触发条件影响。建议使用构建内容标识或随接口变化更新运行时版本，不改解析规则。
- **P2，文件读写完整性**：VmlHost.ets:176、206 单次读写，没有检查短读、短写；发生短读时可能把截断 HTML 当完整文件解析，发生短写时可能仍返回成功。建议循环完成并校验总字节数，具体设备发生概率未测试。
- **P2，宿主重新挂载**：VmlHost.ets:47 销毁时没有清除 ready；复用同一 controller 时，在下一次 onPageEnd 前调用 parseFile 会立即返回失败，而不会等新页面就绪。建议销毁时重置就绪状态，保持现有超时错误报告。
- **解析接受范围扩大**：当前条件标记识别忽略大小写并接受更多写法；shape 的 style 字段顺序不受限制，缺失 z-index 时降级为 0。原始代码对条件标记及字段顺序更严格，要求正则中的 z-index 存在。因此不能把当前更宽松的行为称为逐输入等价。严格基准下应恢复原始筛选条件；无效尺寸处理则需明确保留的是何种原始结果，不能默认“跳过”就是等价。
- **原始辅助函数缺失**：已在 D:\VML 的 JS 源文件中检索，原始文件调用的 toPt、isTableEmpty、getVMLTextAlign 没有找到定义。当前 parseLengthToPt、isEmptyPlaceholderTable、textAlign 回退不能证明与它们等价。特别是空表判断与 td 垂直对齐，需以原始辅助函数源码或原工程测试输出作为后续依据，不应猜测后直接宣称一致。
- **段落转换的扫描边界**：VML_P2DIV 当前通过字符串搜索 p 开闭标签以保护 DOMParser 前的 table 结构，这符合此前的预处理约束；但没有跳过注释和 script/style 文本，里面出现类似 p 的文本时可能被当作真实段落。建议只识别真实标签，并完整跳过注释，不使用 template 重建 VML 注释。尚未进行完整 HTML tokenizer 级验证。

## 本次未判为问题的规则

- 无 wrap 时保留；有 wrap 时按 margin/page 锚点筛选，是原始预期规则，不建议开放“所有 wrap”选项。
- 已有定位 span 删除全部内联 style；其外层直接 span 只删除 position，与原始两个不同操作对应。
- 直接 div > img 跳过；多 shape 的字段分别取最小值；line-height 使用 VML height，不能自行改为包围盒或强制 height。
- replaceVMLSrcs 仍保留 async 声明，但其当前函数体没有 await，源替换同步完成；目前调用它后立即调用 changeVMLSrcs 不存在由这个声明本身造成的时序错误。以后若新增异步 I/O，必须同时改调用链。
- table 在最近 VML 容器之外时不提取为 originalTable，不进行该 table 的提升/删除处理；table 内提升只到 table 同级是用户明确要求，不能按原文件的更高层移动方式回退。

## 验证与产物

HAR 已成功构建：`outlook-vml-har/build/default/outputs/default/outlook_vml_webview.har`。安装本地工程依赖后，entry 的 default@CompileArkTS 也已通过；本次没有执行完整 HAP 打包。

运行时源与根目录 refactored JS 的 SHA-256 均为 `50518F2C84FEA47C3AB57EAA59D18338FE2E276FA0AFBCD777A2B2DC4AE2B735`。

通过 JS 语法检查和生成器同步。可运行 `node outlook-vml-har/tools/audit-fixed-logic.cjs` 复现上述文件名、重复 ID、重复源、li 坐标、字符集引号差异。该脚本只在隔离 VM 中暴露内部函数，没有修改生产 API；它确认问题存在，不代表完整 DOM 差分测试通过。

没有执行真机 ArkWeb 显示验证，也没有因同步 HAR 而改写审核列出的问题。新 HAR 是去配置化同步产物，不是上述问题的修复版。

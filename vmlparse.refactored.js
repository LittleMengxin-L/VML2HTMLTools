/**
 * Outlook 经典版 VML 图片转换器（重构版）
 *
 * 目标：把 Outlook 条件注释中的 VML shape 信息映射到普通 img，
 * 为不支持 VML 的 WebView 提供可用的绝对定位信息。
 *
 * 特点：
 * - 不依赖 jQuery；但入口兼容 jQuery 对象、Element 和 DocumentFragment。
 * - 按 Extract -> Normalize -> Plan -> Apply 四个阶段组织。
 * - 精确匹配 v:shapes ID，支持常用 CSS 长度单位。
 * - 严格按原 vmlparse.js 的固定规则处理图片结构和占位 table。
 * - 所有入口返回 stats/issues/errors/warnings，便于定位特定邮件兼容问题。
 *
 * 两条主要业务链路：
 *
 * 1. 回复/转发/再次编辑已有 VML 的邮件：
 *    const report = OutlookVMLParser.changeVMLSrcs($content);
 *
 * 2. 阅读邮件并在 WebView 中展示 VML 图片：
 *    html = OutlookVMLParser.VML_P2DIV(html); // 注入 WebView 前执行
 *    // WebView 完成 DOM 创建后：
 *    const report = OutlookVMLParser.parseVMLImages($content);
 *
 * VML_P2DIV 是展示链路的结构预处理函数：将承载 VML 的 p 替换为 div，
 * 防止 Outlook 生成的 p > span > table 结构被 p 的内容模型/段落边界截断。
 */
(function attachOutlookVMLParser(global) {
    'use strict';

    const VERSION = '1.0.0';

    const PT_PER_UNIT = Object.freeze({
        pt: 1,
        px: 72 / 96,
        in: 72,
        cm: 72 / 2.54,
        mm: 72 / 25.4,
        pc: 12
    });

    /**
     * 错误码是稳定 API；message/cause/action 可调整，业务判断只应依赖 code。
     * I=输入，V=VML解析，M=关联/布局，D=DOM写入，S=资源路径，P=段落转换，X=内部异常。
     */
    const ERROR_CODES = Object.freeze({
        I1001: Object.freeze({ severity: 'error', stage: 'input', message: '根节点无效。', cause: '入口参数不是 Element、DocumentFragment 或有效 jQuery 对象。', action: '终止本次处理，不修改 DOM。' }),
        I1002: Object.freeze({ severity: 'error', stage: 'input', message: '根节点缺少必要的 DOM 能力。', cause: '节点不支持 querySelectorAll、innerHTML 或属性操作。', action: '终止本次处理。' }),
        I1003: Object.freeze({ severity: 'error', stage: 'input', message: 'HTML 参数类型无效。', cause: 'VML_P2DIV 收到的值不是字符串。', action: '返回空字符串或原值，不执行转换。' }),
        I1008: Object.freeze({ severity: 'info', stage: 'input', message: '输入内容为空。', cause: '根节点没有 HTML 内容。', action: '安全结束，不修改 DOM。' }),
        I1009: Object.freeze({ severity: 'warning', stage: 'input', message: '外部诊断报告对象无效。', cause: 'suppliedReport 缺少 stats/issues/errors/warnings/infos 数组。', action: '创建新的诊断报告。' }),

        V2001: Object.freeze({ severity: 'info', stage: 'extract', message: '未发现 VML 条件注释。', cause: '内容不含受支持的 Outlook VML 条件块。', action: '安全结束，保留原内容。' }),
        V2002: Object.freeze({ severity: 'warning', stage: 'extract', message: '未找到可用的 VML 定位容器。', cause: 'VML 不在 div、li、h2 候选容器内，或容器被规则排除。', action: '不执行图片定位。' }),
        V2003: Object.freeze({ severity: 'warning', stage: 'extract', message: 'VML shape 块格式不完整。', cause: '无法找到合法的 v:shape 开始标签或闭合边界。', action: '跳过该 shape。' }),
        V2004: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML shape 缺少 ID。', cause: 'v:shape 没有 id，或 id 解码后为空。', action: '跳过该 shape，无法与 img 关联。' }),
        V2005: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML shape 缺少 style。', cause: 'v:shape 未提供内联 style。', action: '保留记录，但通常因缺少几何信息而跳过。' }),
        V2006: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML 长度缺失或无法解析。', cause: '坐标/尺寸为空、格式非法或使用了不支持的单位。', action: '该字段记为 null；几何信息不完整时跳过定位。' }),
        V2007: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML shape 尺寸无效。', cause: 'width 或 height 小于等于 0。', action: '跳过该 shape 的布局计算。' }),
        V2008: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML z-index 无效。', cause: 'z-index 不是整数。', action: '使用 0 作为降级值。' }),
        V2009: Object.freeze({ severity: 'warning', stage: 'index', message: '发现重复的 VML shape ID。', cause: '多个 shape 使用同一个 id。', action: '保留首次出现的 shape。' }),
        V2010: Object.freeze({ severity: 'info', stage: 'extract', message: 'VML shape 被环绕规则过滤。', cause: 'w:wrap 锚点不符合当前兼容策略。', action: '不对该 shape 执行绝对定位。' }),
        V2011: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML shape 缺少 imagedata。', cause: 'shape 内没有 v:imagedata 标签。', action: '仍可用于位置关联，但不执行该 shape 的源路径替换。' }),
        V2012: Object.freeze({ severity: 'warning', stage: 'normalize', message: 'VML imagedata 缺少 src。', cause: 'v:imagedata 没有 src 或 src 为空。', action: '仅在普通 img 可提供文件名时尝试补写。' }),

        M3001: Object.freeze({ severity: 'warning', stage: 'associate', message: '图片的 v:shapes 为空。', cause: 'img 存在 v:shapes 属性但未包含有效 ID。', action: '保留原图并跳过定位。' }),
        M3002: Object.freeze({ severity: 'warning', stage: 'associate', message: '未找到图片对应的 VML shape。', cause: 'v:shapes 中的 ID 均未出现在 shape 索引。', action: '保留原图并跳过定位。' }),
        M3003: Object.freeze({ severity: 'warning', stage: 'associate', message: '图片只匹配到部分 VML shape。', cause: 'v:shapes 中部分 ID 不存在。', action: '使用已匹配 shape 计算，并记录缺失 ID。' }),
        M3004: Object.freeze({ severity: 'warning', stage: 'plan', message: '无法确定相对定位容器。', cause: '匹配 shape 没有关联有效容器。', action: '保留原图并跳过定位。' }),
        M3005: Object.freeze({ severity: 'warning', stage: 'plan', message: '节点已脱离当前处理根节点。', cause: '图片或容器在生成计划前已被其他逻辑移动/删除。', action: '跳过该图片。' }),
        M3006: Object.freeze({ severity: 'warning', stage: 'plan', message: '组合几何信息无效。', cause: '没有完整、有限且尺寸为正的 shape 包围盒。', action: '保留原图并跳过定位。' }),
        M3008: Object.freeze({ severity: 'info', stage: 'plan', message: '直接位于 div 下的图片按原逻辑跳过。', cause: '原 vmlparse.js 不处理 div > img 结构。', action: '保持旧版环绕图片结构。' }),
        M3010: Object.freeze({ severity: 'info', stage: 'plan', message: '链接图片缺少外层 span。', cause: 'img 的直接父节点是 a，但在最近 VML 容器内没有找到 span。', action: '按照原 vmlparse.js 行为跳过该图片。' }),

        D4001: Object.freeze({ severity: 'error', stage: 'apply', message: '图片节点没有父节点。', cause: '无法创建包装节点，图片可能已被移除。', action: '跳过该图片。' }),
        D4002: Object.freeze({ severity: 'error', stage: 'apply', message: '缺少 ownerDocument。', cause: '节点不属于可创建元素的 Document。', action: '跳过该图片。' }),
        D4003: Object.freeze({ severity: 'error', stage: 'apply', message: '定位目标无效。', cause: '策略未能产生可设置样式的 Element。', action: '跳过该图片。' }),
        D4004: Object.freeze({ severity: 'error', stage: 'apply', message: '检测到 DOM 循环移动风险。', cause: '目标容器位于待移动节点内部。', action: '禁止移动并跳过。' }),
        D4005: Object.freeze({ severity: 'error', stage: 'apply', message: '应用 VML 布局失败。', cause: 'DOM 包裹、样式写入、节点移动或属性写入抛出异常。', action: '保留可恢复的现有结构并记录原始异常。' }),
        D4006: Object.freeze({ severity: 'info', stage: 'apply', message: '占位 table 未删除。', cause: 'table 仍包含可见或业务节点，不满足安全删除条件。', action: '保留 table。' }),

        S5001: Object.freeze({ severity: 'warning', stage: 'source', message: '普通图片 src 无效。', cause: 'src 为空或无法提取文件名。', action: '不使用该图片更新 VML 源路径。' }),
        S5002: Object.freeze({ severity: 'warning', stage: 'source', message: 'VML 图片源无法替换。', cause: 'shape ID、imagedata 或目标文件名缺失。', action: '保留原 VML src。' }),
        S5003: Object.freeze({ severity: 'error', stage: 'source', message: 'WebView 图片 src 替换为 VML 图片 src 报错。', cause: 'changeVMLSrcs 从普通 img.src 向 VML imagedata src 同步时发生异常。', action: '终止本次 VML 图片源改写并记录原始异常。' }),
        S5004: Object.freeze({ severity: 'warning', stage: 'source', message: '同一 shape ID 对应多个普通图片源。', cause: '多个 img 使用相同 v:shapes ID，但文件名不同。', action: '保留首次出现的图片源。' }),
        S5005: Object.freeze({ severity: 'warning', stage: 'source', message: 'VML 图片 src 替换为 WebView 图片 src 报错。', cause: 'replaceVMLSrcs 从 VML imagedata src 向普通 img.src 同步时发生异常。', action: '保留 WebView 图片原 src，并继续执行 VML 布局解析。' }),

        P6002: Object.freeze({ severity: 'error', stage: 'paragraph', message: 'VML 段落预处理失败。', cause: '扫描或改写原始 HTML 字符串时发生异常。', action: '返回原 HTML。' }),
        P6003: Object.freeze({ severity: 'warning', stage: 'paragraph', message: 'VML 段落结构不完整。', cause: '含 VML 的 p 缺少合法开始标签边界或闭合标签。', action: '保留未完成的原始内容，避免错误截取后续邮件。' }),

        X9001: Object.freeze({ severity: 'error', stage: 'internal', message: '未捕获的内部异常。', cause: '入口保护捕获到非预期错误。', action: '停止当前入口，返回诊断报告。' })
    });

    function createReport() {
        return {
            stats: {
                shapes: 0,
                images: 0,
                matched: 0,
                transformed: 0,
                skipped: 0,
                sourcesRewritten: 0,
                paragraphsConverted: 0
            },
            issues: [],
            errors: [],
            warnings: [],
            infos: [],
            status: 'ok'
        };
    }

    function addIssue(report, code, details) {
        const definition = ERROR_CODES[code] || ERROR_CODES.X9001;
        const issue = {
            code,
            severity: definition.severity,
            stage: definition.stage,
            message: definition.message,
            cause: definition.cause,
            action: definition.action,
            ...(details || {})
        };
        report.issues.push(issue);
        if (definition.severity === 'error') report.errors.push(issue);
        else if (definition.severity === 'warning') report.warnings.push(issue);
        else report.infos.push(issue);
        report.status = report.errors.length ? 'failed' : (report.warnings.length ? 'partial' : 'ok');
        return issue;
    }

    function resolveReport(suppliedReport) {
        if (suppliedReport == null) return createReport();
        try {
            const statKeys = [
                'shapes', 'images', 'matched', 'transformed', 'skipped',
                'sourcesRewritten', 'paragraphsConverted'
            ];
            const valid = suppliedReport.stats &&
                statKeys.every(key => Number.isFinite(suppliedReport.stats[key])) &&
                ['issues', 'errors', 'warnings', 'infos'].every(key =>
                    Array.isArray(suppliedReport[key]) && !Object.isFrozen(suppliedReport[key])
                ) && !Object.isFrozen(suppliedReport);
            if (valid) return suppliedReport;
        } catch (_) {
            // Proxy getter 等异常统一按无效报告对象降级。
        }
        const report = createReport();
        addIssue(report, 'I1009');
        return report;
    }

    function resolveRoot(input) {
        if (!input) return null;
        if (input.nodeType === 1 || input.nodeType === 11) return input;
        if (input[0] && (input[0].nodeType === 1 || input[0].nodeType === 11)) {
            return input[0];
        }
        return null;
    }

    function validateRoot(input, report) {
        const root = resolveRoot(input);
        if (!root) {
            addIssue(report, 'I1001', { receivedType: input == null ? String(input) : typeof input });
            return null;
        }
        if (typeof root.querySelectorAll !== 'function' || !('innerHTML' in root)) {
            addIssue(report, 'I1002');
            return null;
        }
        if (!root.innerHTML) addIssue(report, 'I1008');
        return root;
    }

    function decodeNumericEntities(value) {
        if (!value) return value;
        return String(value).replace(/&#(?:x([0-9a-f]+)|(\d+));/gi, (match, hex, decimal) => {
            const codePoint = Number.parseInt(hex || decimal, hex ? 16 : 10);
            if (!Number.isFinite(codePoint) || codePoint < 0 || codePoint > 0x10ffff) {
                return match;
            }
            try {
                return String.fromCodePoint(codePoint);
            } catch (_) {
                return match;
            }
        });
    }

    function escapeRegExp(value) {
        return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function escapeAttribute(value) {
        return String(value)
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }

    function readAttribute(tag, name) {
        const attributeName = escapeRegExp(name);
        const match = String(tag).match(
            new RegExp(`(?:^|\\s)${attributeName}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, 'i')
        );
        return match ? (match[1] ?? match[2] ?? match[3] ?? '') : null;
    }

    function replaceAttribute(tag, name, value) {
        const attributeName = escapeRegExp(name);
        const expression = new RegExp(
            `((?:^|\\s)${attributeName}\\s*=\\s*)(?:"[^"]*"|'[^']*'|[^\\s>]+)`,
            'i'
        );
        if (!expression.test(tag)) return tag;
        return tag.replace(expression, `$1"${escapeAttribute(value)}"`);
    }

    function upsertAttribute(tag, name, value) {
        if (readAttribute(tag, name) != null) return replaceAttribute(tag, name, value);
        return String(tag).replace(/(\/?>)\s*$/, ` ${name}="${escapeAttribute(value)}"$1`);
    }

    function parseStyle(styleText) {
        const declarations = new Map();
        if (!styleText) return declarations;

        String(styleText).split(';').forEach(declaration => {
            const separator = declaration.indexOf(':');
            if (separator < 1) return;
            const name = declaration.slice(0, separator).trim().toLowerCase();
            const value = declaration.slice(separator + 1).trim();
            if (name) declarations.set(name, value);
        });
        return declarations;
    }

    function parseLengthToPt(rawValue) {
        if (typeof rawValue === 'number') {
            return Number.isFinite(rawValue) ? rawValue : null;
        }
        if (rawValue == null) return null;

        const match = String(rawValue).trim().match(
            /^(-?(?:\d+(?:\.\d+)?|\.\d+))(pt|px|in|cm|mm|pc)?$/i
        );
        if (!match) return null;

        const number = Number.parseFloat(match[1]);
        const unit = (match[2] || 'pt').toLowerCase();
        const factor = PT_PER_UNIT[unit];
        return Number.isFinite(number) && factor ? number * factor : null;
    }

    function parseZIndex(rawValue) {
        if (rawValue == null || rawValue === '') return 0;
        const parsed = Number.parseInt(String(rawValue).trim(), 10);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function getResourceFileName(src) {
        if (!src) return null;
        const text = String(src).trim();
        if (!text || text.length > 2048 || /[\u0000-\u001f\u007f]/.test(text)) return null;
        if (/^(?:data|blob|javascript):/i.test(text)) return null;
        const clean = text.replace(/^cid:/i, '').split(/[?#]/, 1)[0].replace(/\\/g, '/');
        const fileName = clean.slice(clean.lastIndexOf('/') + 1);
        if (!fileName) return null;
        try {
            const decoded = decodeURIComponent(fileName);
            return /[/\\\u0000-\u001f\u007f]/.test(decoded) ? fileName : decoded;
        } catch (_) {
            return fileName;
        }
    }

    function splitShapeIds(value) {
        if (!value) return [];
        return [...new Set(
            decodeNumericEntities(String(value))
                .trim()
                .split(/\s+/)
                .filter(Boolean)
        )];
    }

    function containsVmlCondition(html) {
        return /<!--\s*\[if\s+(?:gte\s+)?vml\b/i.test(html || '');
    }

    /**
     * 只使用正则定位原始 VML 字符串边界；属性和 style 分别解析，
     * 避免依赖 style 属性顺序。
     */
    function extractShapeBlocks(html) {
        const source = String(html || '');
        const blocks = [];
        const fullShape = /<v:shape\b[\s\S]*?<\/v:shape\s*>/gi;
        let match;

        while ((match = fullShape.exec(source)) !== null) {
            blocks.push(match[0]);
        }

        // 兼容极少数自闭合 shape；避免重复提取完整 shape 的开始标签。
        const sourceWithoutFullShapes = source.replace(fullShape, '');
        const selfClosingShape = /<v:shape\b[^>]*\/\s*>/gi;
        while ((match = selfClosingShape.exec(sourceWithoutFullShapes)) !== null) {
            blocks.push(match[0]);
        }
        return blocks;
    }

    function parseShape(block, container, report) {
        const startTagMatch = String(block).match(/^<v:shape\b[^>]*>/i);
        if (!startTagMatch) {
            addIssue(report, 'V2003', { rawPreview: String(block).slice(0, 160) });
            return null;
        }

        const startTag = startTagMatch[0];
        const rawId = readAttribute(startTag, 'id');
        const id = decodeNumericEntities(rawId || '');
        if (!id) {
            addIssue(report, 'V2004', { rawPreview: startTag.slice(0, 160) });
            return null;
        }

        const styleText = readAttribute(startTag, 'style');
        const style = parseStyle(styleText);
        if (!styleText) {
            addIssue(report, 'V2005', { shapeId: id });
        }

        const rawGeometry = {
            left: style.get('margin-left'),
            top: style.get('margin-top'),
            width: style.get('width'),
            height: style.get('height')
        };
        const geometry = Object.fromEntries(Object.entries(rawGeometry).map(([field, rawValue]) => {
            const parsed = parseLengthToPt(rawValue);
            if (parsed == null) addIssue(report, 'V2006', { shapeId: id, field, rawValue: rawValue ?? null });
            return [field, parsed];
        }));
        if ((geometry.width != null && geometry.width <= 0) || (geometry.height != null && geometry.height <= 0)) {
            addIssue(report, 'V2007', {
                shapeId: id,
                width: geometry.width,
                height: geometry.height
            });
            geometry.width = null;
            geometry.height = null;
        }

        const rawZIndex = style.get('z-index');
        if (rawZIndex != null && !/^-?\d+$/.test(String(rawZIndex).trim())) {
            addIssue(report, 'V2008', { shapeId: id, rawValue: rawZIndex });
        }

        const imageTag = String(block).match(/<v:imagedata\b[^>]*>/i)?.[0] || null;
        const wrapTag = String(block).match(/<w:wrap\b[^>]*>/i)?.[0] || null;
        const imageSrc = imageTag ? readAttribute(imageTag, 'src') : null;
        if (!imageTag) addIssue(report, 'V2011', { shapeId: id });
        else if (!imageSrc) addIssue(report, 'V2012', { shapeId: id });

        return {
            id,
            imageSrc,
            boxPt: geometry,
            zIndex: parseZIndex(rawZIndex),
            textAlign: style.get('text-align')?.toLowerCase() || null,
            anchor: {
                x: wrapTag ? readAttribute(wrapTag, 'anchorx')?.toLowerCase() || null : null,
                y: wrapTag ? readAttribute(wrapTag, 'anchory')?.toLowerCase() || null : null
            },
            wrapType: wrapTag ? readAttribute(wrapTag, 'type')?.toLowerCase() || null : null,
            hasWrap: Boolean(wrapTag),
            container,
            raw: block
        };
    }

    function isShapeEligible(shape) {
        if (!shape.hasWrap) return true;
        // 保持旧实现的筛选语义：浮于文字上/衬于文字下相关锚点进入定位流程。
        return shape.anchor.x === 'margin' || shape.anchor.y === 'page';
    }

    function isExcludedContainer(element) {
        return element.classList.contains('WordSection1') ||
            element.id === 'mail-content' ||
            element.id === 'mail-header';
    }

    function isNearestVmlContainer(element) {
        // 严格对应 vmlparse.js：只要任一直接子元素内部仍含 VML，
        // 当前元素就不是距离该 VML 最近的候选容器。
        return !Array.from(element.children).some(child =>
            containsVmlCondition(child.innerHTML)
        );
    }

    function extractShapes(root, report) {
        const records = [];
        const candidates = Array.from(root.querySelectorAll('div, li, h2'));
        if (!containsVmlCondition(root.innerHTML)) {
            addIssue(report, 'V2001');
            return records;
        }
        let eligibleContainerCount = 0;

        candidates.forEach(container => {
            if (isExcludedContainer(container)) return;
            if (!containsVmlCondition(container.innerHTML)) return;
            if (!isNearestVmlContainer(container)) return;
            eligibleContainerCount += 1;

            const blocks = extractShapeBlocks(container.innerHTML);
            if (!blocks.length) addIssue(report, 'V2003', { containerTag: container.tagName.toLowerCase() });
            blocks.forEach(block => {
                const shape = parseShape(block, container, report);
                if (!shape) return;
                if (isShapeEligible(shape)) records.push(shape);
                else addIssue(report, 'V2010', { shapeId: shape.id, anchor: shape.anchor, wrapType: shape.wrapType });
            });
        });

        if (!eligibleContainerCount) addIssue(report, 'V2002');
        report.stats.shapes = records.length;
        return records;
    }

    function buildShapeIndex(shapes, report) {
        const shapeById = new Map();
        shapes.forEach(shape => {
            // Outlook HTML 中偶尔出现重复 ID。保留第一次出现项，避免后续父容器覆盖最近容器。
            if (!shapeById.has(shape.id)) shapeById.set(shape.id, shape);
            else addIssue(report, 'V2009', { shapeId: shape.id });
        });
        return shapeById;
    }

    function combineShapeGeometry(shapes) {
        const valid = shapes.filter(shape => {
            const box = shape.boxPt;
            return [box.left, box.top, box.width, box.height].every(Number.isFinite);
        });
        if (!valid.length) return null;

        const left = Math.min(...valid.map(shape => shape.boxPt.left));
        const top = Math.min(...valid.map(shape => shape.boxPt.top));
        return {
            left,
            top,
            // 原实现对多个 shape 的每个字段分别取最小值，而不是计算包围盒。
            width: Math.min(...valid.map(shape => shape.boxPt.width)),
            height: Math.min(...valid.map(shape => shape.boxPt.height)),
            zIndex: Math.min(...valid.map(shape => shape.zIndex))
        };
    }

    function readContainerOffset(container) {
        const style = parseStyle(container.getAttribute('style'));
        return {
            left: (parseLengthToPt(style.get('margin-left')) || 0) +
                (parseLengthToPt(style.get('text-indent')) || 0),
            top: parseLengthToPt(style.get('margin-top')) || 0,
            textAlign: style.get('text-align')?.toLowerCase() ||
                container.getAttribute('align')?.toLowerCase() || null
        };
    }

    function computeLayout(shapes, container) {
        const box = combineShapeGeometry(shapes);
        if (!box) return null;

        const offset = readContainerOffset(container);
        const isListItem = container.tagName.toLowerCase() === 'li';
        return {
            left: isListItem ? box.left : box.left - offset.left,
            top: isListItem ? box.top : box.top - offset.top,
            width: box.width,
            height: box.height,
            zIndex: box.zIndex,
            textAlign: shapes.find(shape => shape.textAlign)?.textAlign || offset.textAlign
        };
    }

    function chooseContainer(shapes, image) {
        const containers = [...new Set(shapes.map(shape => shape.container).filter(Boolean))];
        if (!containers.length) return null;

        // 优先选择真正包含普通 img 的 VML 容器，否则使用 DOM 层级最深的候选项。
        const containingImage = containers.filter(container => container.contains(image));
        const pool = containingImage.length ? containingImage : containers;
        return pool.reduce((deepest, current) => {
            if (!deepest) return current;
            return deepest.contains(current) ? current : deepest;
        }, null);
    }

    function findAncestorWithin(start, selector, boundary) {
        let current = start;
        while (current && current !== boundary) {
            if (current.nodeType === 1 && current.matches(selector)) return current;
            current = current.parentElement;
        }
        return null;
    }

    function findClosestTableWithin(image, container) {
        const closestTable = typeof image.closest === 'function' ? image.closest('table') : null;
        return closestTable && container.contains(closestTable) ? closestTable : null;
    }

    function createImagePlan(root, image, shapeById, report) {
        const ids = splitShapeIds(image.getAttribute('v:shapes'));
        if (!ids.length) {
            addIssue(report, 'M3001');
            return { image, action: 'skip', reason: 'M3001' };
        }

        if (!root.contains(image) || !image.parentNode) {
            addIssue(report, 'M3005', { shapeIds: ids, node: 'image' });
            return { image, action: 'skip', reason: 'M3005' };
        }

        const shapes = ids.map(id => shapeById.get(id)).filter(Boolean);
        if (!shapes.length) {
            addIssue(report, 'M3002', { shapeIds: ids });
            return { image, action: 'skip', reason: 'M3002' };
        }
        const matchedIds = new Set(shapes.map(shape => shape.id));
        const missingShapeIds = ids.filter(id => !matchedIds.has(id));
        if (missingShapeIds.length) addIssue(report, 'M3003', { shapeIds: ids, missingShapeIds });

        const container = chooseContainer(shapes, image);
        if (!container) {
            addIssue(report, 'M3004', { shapeIds: ids });
            return { image, action: 'skip', reason: 'M3004' };
        }
        if (!root.contains(container)) {
            addIssue(report, 'M3005', { shapeIds: ids, node: 'container' });
            return { image, action: 'skip', reason: 'M3005' };
        }

        const layout = computeLayout(shapes, container);
        if (!layout) {
            addIssue(report, 'M3006', { shapeIds: ids });
            return { image, action: 'skip', reason: 'M3006' };
        }
        if (image.parentElement?.tagName.toLowerCase() === 'div') {
            addIssue(report, 'M3008', { shapeIds: ids });
            return { image, action: 'skip', reason: 'M3008' };
        }

        let target = null;
        let strategy = 'new-wrapper';
        const directParent = image.parentElement;
        const originalTable = findClosestTableWithin(image, container);

        if (directParent?.tagName.toLowerCase() === 'span') {
            target = directParent;
            strategy = 'existing-span';
        } else if (directParent?.tagName.toLowerCase() === 'a') {
            target = findAncestorWithin(directParent.parentElement, 'span', container);
            if (!target) {
                addIssue(report, 'M3010', { shapeIds: ids });
                return { image, action: 'skip', reason: 'M3010' };
            }
            if (originalTable) {
                target = directParent;
                strategy = 'link-table-wrapper';
            } else {
                strategy = 'link-existing-span';
            }
        }

        return {
            image,
            ids,
            shapes,
            container,
            originalTable,
            layout,
            target,
            strategy,
            action: 'transform'
        };
    }

    function buildPlans(root, shapeById, report) {
        const images = Array.from(root.querySelectorAll('img')).filter(image =>
            image.hasAttribute('v:shapes')
        );
        report.stats.images = images.length;

        return images.map(image =>
            createImagePlan(root, image, shapeById, report)
        ).filter(Boolean);
    }

    function ensurePositioningContainer(container) {
        const position = container.style.getPropertyValue('position').toLowerCase();
        if (position !== 'relative') {
            container.style.setProperty('position', 'relative');
        }
    }

    function setLayoutStyles(element, layout) {
        element.style.setProperty('position', 'absolute');
        element.style.setProperty('left', `${layout.left}pt`);
        element.style.setProperty('top', `${layout.top}pt`);
        element.style.setProperty('width', `${layout.width}pt`);
        element.style.setProperty('line-height', `${layout.height}pt`);
        element.style.setProperty('z-index', String(layout.zIndex));
    }

    function removePositionFromParentSpan(element) {
        const parent = element.parentElement;
        if (!parent || parent.tagName.toLowerCase() !== 'span') return;
        parent.style.removeProperty('position');
        if (!parent.getAttribute('style')?.trim()) parent.removeAttribute('style');
    }

    function createCodedError(code, message) {
        const error = new Error(message || ERROR_CODES[code]?.message || code);
        error.vmlCode = code;
        return error;
    }

    function wrapElement(element) {
        if (!element?.parentNode) throw createCodedError('D4001');
        if (!element.ownerDocument || typeof element.ownerDocument.createElement !== 'function') {
            throw createCodedError('D4002');
        }
        const wrapper = element.ownerDocument.createElement('span');
        element.parentNode.insertBefore(wrapper, element);
        wrapper.appendChild(element);
        return wrapper;
    }

    function isIgnorableNode(node, controlledNode) {
        if (node === controlledNode || node.contains?.(controlledNode)) return true;
        if (node.nodeType === 3) return !node.nodeValue.trim();
        if (node.nodeType === 8) return true;
        if (node.nodeType !== 1) return true;
        if (['img', 'video', 'audio', 'canvas', 'svg', 'iframe'].includes(node.tagName.toLowerCase())) {
            return false;
        }
        return Array.from(node.childNodes).every(child => isIgnorableNode(child, controlledNode));
    }

    function isEmptyPlaceholderTable(table, controlledNode) {
        if (!table) return false;
        return Array.from(table.childNodes).every(node => isIgnorableNode(node, controlledNode));
    }

    function updateCellAlignment(container, layout) {
        const cell = container.parentElement;
        if (!cell || cell.tagName.toLowerCase() !== 'td') return;
        cell.style.setProperty('vertical-align', layout.textAlign === 'center' ? 'middle' : 'top');
    }

    function applyPlan(plan, report) {
        if (plan.action !== 'transform') {
            report.stats.skipped += 1;
            return;
        }

        const { image, container, layout, originalTable } = plan;

        try {
            if (!image.parentNode) throw createCodedError('D4001');
            if (!image.ownerDocument) throw createCodedError('D4002');
            if (!container?.style || typeof container.contains !== 'function') {
                throw createCodedError('D4003');
            }
            ensurePositioningContainer(container);
            updateCellAlignment(container, layout);

            let positionedNode = plan.target;
            if (plan.strategy === 'new-wrapper') {
                positionedNode = wrapElement(image);
            } else if (plan.strategy === 'link-table-wrapper') {
                positionedNode = wrapElement(plan.target);
            }

            if (!positionedNode?.style || typeof positionedNode.setAttribute !== 'function') {
                throw createCodedError('D4003');
            }
            if (positionedNode.contains(container)) throw createCodedError('D4004');

            // 严格对应 vmlparse.js：已有 span 先删除全部内联样式；
            // 新建包装 span 本身没有需要清理的旧样式。
            if (plan.strategy === 'existing-span' || plan.strategy === 'link-existing-span') {
                positionedNode.removeAttribute('style');
            }
            setLayoutStyles(positionedNode, layout);
            if (plan.strategy === 'existing-span' || plan.strategy === 'link-existing-span') {
                removePositionFromParentSpan(positionedNode);
            }
            image.classList.add('vml-controlled');

            // 图片仍被 Outlook 占位 table 包裹时，只提升到 table 同级，
            // 保留 table 之上的 span/div 等邮件原始层级。
            const isInsideOriginalTable = originalTable?.contains(positionedNode) === true;
            if (isInsideOriginalTable && originalTable.parentNode) {
                originalTable.parentNode.insertBefore(positionedNode, originalTable);
            } else if (!container.contains(positionedNode)) {
                container.appendChild(positionedNode);
            }

            if (originalTable && originalTable.isConnected) {
                const tableIsEmpty = isEmptyPlaceholderTable(originalTable, positionedNode);
                const legacyHierarchyMatches =
                    container.parentElement === originalTable.parentElement ||
                    container.children[0] === originalTable.parentElement;
                const canRemove = tableIsEmpty &&
                    (plan.strategy === 'link-table-wrapper' ||
                        (plan.strategy === 'new-wrapper' && legacyHierarchyMatches));
                if (canRemove) {
                    originalTable.remove();
                } else {
                    addIssue(report, 'D4006', {
                        shapeIds: plan.ids,
                        reason: !tableIsEmpty ? 'table-not-empty' : 'legacy-hierarchy-not-matched'
                    });
                }
            }

            report.stats.transformed += 1;
        } catch (error) {
            report.stats.skipped += 1;
            addIssue(report, error?.vmlCode || 'D4005', {
                shapeIds: plan.ids,
                originalError: error?.message || String(error)
            });
        }
    }

    function rewriteVmlBlock(block, imageSourceByShapeId, report) {
        const startTagMatch = String(block).match(/^<v:shape\b[^>]*>/i);
        if (!startTagMatch) {
            addIssue(report, 'S5002', { reason: 'shape-start-tag-missing' });
            return { html: block, changed: false };
        }

        const startTag = startTagMatch[0];
        const rawId = readAttribute(startTag, 'id');
        const id = decodeNumericEntities(rawId || '');
        if (!id) {
            addIssue(report, 'S5002', { reason: 'shape-id-missing' });
            return { html: block, changed: false };
        }

        let rewritten = block;
        let changed = false;
        if (rawId !== id) {
            rewritten = rewritten.replace(startTag, replaceAttribute(startTag, 'id', id));
            changed = true;
        }

        const imageTagMatch = rewritten.match(/<v:imagedata\b[^>]*>/i);
        if (!imageTagMatch) {
            addIssue(report, 'S5002', { shapeId: id, reason: 'imagedata-missing' });
            return { html: rewritten, changed };
        }

        const imageTag = imageTagMatch[0];
        const oldSrc = readAttribute(imageTag, 'src');
        const targetSrc = imageSourceByShapeId.get(id) || getResourceFileName(oldSrc);
        if (!targetSrc) {
            addIssue(report, 'S5002', { shapeId: id, reason: 'target-src-missing' });
        } else if (targetSrc !== oldSrc) {
            rewritten = rewritten.replace(imageTag, upsertAttribute(imageTag, 'src', targetSrc));
            changed = true;
        }
        return { html: rewritten, changed };
    }

    function rewriteVmlImageSources(root, report) {
        try {
            const imageSourceByShapeId = new Map();
            Array.from(root.querySelectorAll('img')).forEach(image => {
                if (!image.hasAttribute('v:shapes')) return;
                const ids = splitShapeIds(image.getAttribute('v:shapes'));
                const fileName = getResourceFileName(image.getAttribute('src'));
                if (!fileName) {
                    addIssue(report, 'S5001', { shapeIds: ids });
                    return;
                }
                ids.forEach(id => {
                    if (!imageSourceByShapeId.has(id)) imageSourceByShapeId.set(id, fileName);
                    else if (imageSourceByShapeId.get(id) !== fileName) {
                        addIssue(report, 'S5004', {
                            shapeId: id,
                            keptSource: imageSourceByShapeId.get(id),
                            ignoredSource: fileName
                        });
                    }
                });
            });

            const originalHtml = root.innerHTML;
            const fullShape = /<v:shape\b[\s\S]*?<\/v:shape\s*>/gi;
            const rewrittenHtml = originalHtml.replace(fullShape, block => {
                const result = rewriteVmlBlock(block, imageSourceByShapeId, report);
                if (result.changed) report.stats.sourcesRewritten += 1;
                return result.html;
            });

            // 条件注释中的 VML 只能通过字符串改写；仅在内容确实变化时重建 DOM。
            if (rewrittenHtml !== originalHtml) root.innerHTML = rewrittenHtml;
        } catch (error) {
            addIssue(report, 'S5003', { originalError: error?.message || String(error) });
        }
        return report;
    }

    /**
     * 阅读展示链路的图片源同步，严格对应原始 replaceVMLSrcs：
     * VML imagedata src -> 普通 img src，不修改 VML 条件注释。
     */
    async function replaceVMLSrcs(input, suppliedReport) {
        const report = resolveReport(suppliedReport);
        try {
            const root = validateRoot(input, report);
            if (!root || !root.innerHTML) return report;
            const html = root.innerHTML;

            function decodeUnicodeEntities(value) {
                if (!value) return value;
                return String(value).replace(/&#(\d+);/g, (_match, code) =>
                    String.fromCharCode(Number.parseInt(code, 10))
                );
            }

            let newHtml = html;
            newHtml = newHtml.replace(
                /<v:shape\b(?:(?!<\/v:shape>).)*?\bid="([^"]+)"(?:(?!<\/v:shape>).)*?<v:imagedata[^>]*\bsrc="([^"]+)"/gis,
                (match, id) => {
                    const decodedId = decodeUnicodeEntities(id);
                    return match.replace(`id="${id}"`, `id="${decodedId}"`);
                }
            );

            const vmlShapes = [...newHtml.matchAll(
                /<v:shape\b(?:(?!<\/v:shape>).)*?\bid="([^"]+)"(?:(?!<\/v:shape>).)*?<v:imagedata[^>]*\bsrc="([^"]+)"/gis
            )].map(match => ({ id: match[1], src: match[2] }));
            const images = Array.from(root.querySelectorAll('img'));
            const isImage = source => /\.(png|jpe?g|gif|bmp|webp|svg)$/i.test(source);
            const extractFolderId = source => {
                if (!source) return null;
                try {
                    const path = source.replace(/^file:\/\//, '');
                    const parts = path.split('/');
                    const inlineIndex = parts.indexOf('inline');
                    return inlineIndex > 0 ? parts[inlineIndex - 1] : null;
                } catch (_) {
                    return null;
                }
            };
            const isFolderMatch = (first, second) => {
                if (!first || !second) return false;
                const firstNumber = Number.parseInt(first, 10);
                const secondNumber = Number.parseInt(second, 10);
                if (Number.isNaN(firstNumber) || Number.isNaN(secondNumber)) return false;
                return firstNumber === secondNumber;
            };

            images.forEach(image => {
                const imageShapeId = image.getAttribute('v:shapes');
                if (!imageShapeId) return;
                const shape = vmlShapes.find(item => item.id === imageShapeId);
                if (!shape) return;
                if (isImage(shape.src)) {
                    const imageFolderId = extractFolderId(image.getAttribute('src'));
                    const shapeFolderId = extractFolderId(shape.src);
                    if (isFolderMatch(imageFolderId, shapeFolderId)) {
                        image.setAttribute('src', shape.src);
                    }
                }
            });
        } catch (error) {
            addIssue(report, 'S5005', { originalError: error?.message || String(error) });
        }
        return report;
    }

    function findTagEnd(source, startIndex) {
        let quote = null;
        for (let index = startIndex; index < source.length; index += 1) {
            const character = source[index];
            if (quote) {
                if (character === quote) quote = null;
            } else if (character === '"' || character === "'") {
                quote = character;
            } else if (character === '>') {
                return index;
            }
        }
        return -1;
    }

    function convertVmlParagraphs(html, report) {
        const input = typeof html === 'string' ? html : (html == null ? '' : String(html));
        if (typeof html !== 'string') addIssue(report, 'I1003', { receivedType: html == null ? String(html) : typeof html });
        if (!input) {
            addIssue(report, 'I1008');
            return input;
        }

        try {
            const paragraphStart = /<p\b/gi;
            const paragraphEnd = /<\/p\s*>/gi;
            let cursor = 0;
            let output = '';
            let startMatch;

            while ((startMatch = paragraphStart.exec(input)) !== null) {
                const openStart = startMatch.index;
                const openEnd = findTagEnd(input, openStart);
                if (openEnd < 0) {
                    if (containsVmlCondition(input.slice(openStart))) {
                        addIssue(report, 'P6003', { offset: openStart, reason: 'opening-tag-not-closed' });
                    }
                    break;
                }

                paragraphEnd.lastIndex = openEnd + 1;
                const closeMatch = paragraphEnd.exec(input);
                if (!closeMatch) {
                    if (containsVmlCondition(input.slice(openEnd + 1))) {
                        addIssue(report, 'P6003', { offset: openStart, reason: 'closing-tag-missing' });
                    }
                    break;
                }

                const closeEnd = paragraphEnd.lastIndex;
                const originalBlock = input.slice(openStart, closeEnd);
                output += input.slice(cursor, openStart);
                if (containsVmlCondition(originalBlock)) {
                    const openingTag = input.slice(openStart, openEnd + 1).replace(/^<p\b/i, '<div');
                    output += openingTag + input.slice(openEnd + 1, closeMatch.index) + '</div>';
                    report.stats.paragraphsConverted += 1;
                } else {
                    output += originalBlock;
                }

                cursor = closeEnd;
                paragraphStart.lastIndex = closeEnd;
            }

            return output + input.slice(cursor);
        } catch (error) {
            addIssue(report, 'P6002', { originalError: error?.message || String(error) });
            return input;
        }
    }

    /**
     * 编辑链路主函数。
     * 用于 WebView 邮箱回复、转发、再次编辑已有 VML 的邮件，使 VML imagedata
     * 与普通 img 使用一致的本地资源文件名；不负责把 VML 坐标转换为展示布局。
     */
    function changeVMLSrcs(input, suppliedReport) {
        const report = resolveReport(suppliedReport);
        try {
            const root = validateRoot(input, report);
            if (!root || !root.innerHTML) return report;
            return rewriteVmlImageSources(root, report);
        } catch (error) {
            addIssue(report, 'S5003', { entry: 'changeVMLSrcs', originalError: error?.message || String(error) });
            return report;
        }
    }

    /**
     * 展示链路预处理函数，应在 HTML 注入 WebView 之前执行。
     * 将含 VML 条件块的 p 替换为 div，避免 p > span > table 被截断。
     */
    function VML_P2DIV(html, suppliedReport) {
        const report = resolveReport(suppliedReport);
        try {
            return convertVmlParagraphs(html, report);
        } catch (error) {
            addIssue(report, 'X9001', { entry: 'VML_P2DIV', originalError: error?.message || String(error) });
            return typeof html === 'string' ? html : '';
        }
    }

    /**
     * 展示链路主函数，应在 WebView 已生成 DOM 后执行。
     * 解析 VML shape 的位置和尺寸，并将其映射到普通 img/包装节点。
     * 本函数不承担回复、转发或再次编辑时的 VML 源路径修复职责。
     */
    function parseVMLImages(input, suppliedReport) {
        const report = resolveReport(suppliedReport);
        try {
            const root = validateRoot(input, report);
            if (!root || !root.innerHTML) return report;

            replaceVMLSrcs(root, report);
            changeVMLSrcs(root, report);
            const shapes = extractShapes(root, report);
            const shapeById = buildShapeIndex(shapes, report);
            const plans = buildPlans(root, shapeById, report);
            report.stats.matched = plans.filter(plan => plan.action === 'transform').length;
            plans.forEach(plan => applyPlan(plan, report));
            return report;
        } catch (error) {
            addIssue(report, 'X9001', { entry: 'parseVMLImages', originalError: error?.message || String(error) });
            return report;
        }
    }

    const api = Object.freeze({
        VERSION,
        ERROR_CODES,
        createReport,
        changeVMLSrcs,
        replaceVMLSrcs,
        VML_P2DIV,
        parseVMLImages,
        utils: Object.freeze({
            decodeNumericEntities,
            getResourceFileName,
            parseLengthToPt,
            parseStyle,
            splitShapeIds
        })
    });

    global.OutlookVMLParser = api;

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = api;
    }
})(typeof window !== 'undefined' ? window : globalThis);

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
 * - 默认不删除 Outlook 生成的 table，避免误删邮件内容。
 * - 通过 data-vml-processed 保证重复执行不会二次包裹节点。
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
 * transform 只是可选的组合入口，不能替代调用方对编辑链路和展示链路的区分。
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

    const DEFAULT_OPTIONS = Object.freeze({
        rewriteImageSources: true,
        convertVmlParagraphs: false,
        processAllWrapTypes: false,
        removeEmptyPlaceholderTables: false,
        skipDirectDivImages: true,
        markAttribute: 'data-vml-processed',
        maxCoordinatePt: 10000000,
        debug: false
    });

    /**
     * 错误码是稳定 API；message/cause/action 可调整，业务判断只应依赖 code。
     * I=输入，V=VML解析，M=关联/布局，D=DOM写入，S=资源路径，P=段落转换，X=内部异常。
     */
    const ERROR_CODES = Object.freeze({
        I1001: Object.freeze({ severity: 'error', stage: 'input', message: '根节点无效。', cause: '入口参数不是 Element、DocumentFragment 或有效 jQuery 对象。', action: '终止本次处理，不修改 DOM。' }),
        I1002: Object.freeze({ severity: 'error', stage: 'input', message: '根节点缺少必要的 DOM 能力。', cause: '节点不支持 querySelectorAll、innerHTML 或属性操作。', action: '终止本次处理。' }),
        I1003: Object.freeze({ severity: 'error', stage: 'input', message: 'HTML 参数类型无效。', cause: 'VML_P2DIV 收到的值不是字符串。', action: '返回空字符串或原值，不执行转换。' }),
        I1004: Object.freeze({ severity: 'warning', stage: 'input', message: 'options 参数无效。', cause: 'options 不是普通对象。', action: '忽略该参数并使用默认配置。' }),
        I1005: Object.freeze({ severity: 'warning', stage: 'input', message: '布尔配置项类型无效。', cause: '配置值不是 boolean。', action: '该项回退为默认值。' }),
        I1006: Object.freeze({ severity: 'warning', stage: 'input', message: '处理标记属性名无效。', cause: 'markAttribute 不是合法的 data-* 属性名。', action: '回退为 data-vml-processed。' }),
        I1007: Object.freeze({ severity: 'warning', stage: 'input', message: '最大坐标阈值无效。', cause: 'maxCoordinatePt 不是有限正数。', action: '回退为默认阈值。' }),
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
        M3007: Object.freeze({ severity: 'warning', stage: 'plan', message: '布局值超出安全范围。', cause: '坐标、尺寸或计算结果非有限数，或超过 maxCoordinatePt。', action: '保留原图并跳过定位。' }),
        M3008: Object.freeze({ severity: 'info', stage: 'plan', message: '直接位于 div 下的图片按兼容规则跳过。', cause: 'skipDirectDivImages 已启用。', action: '保持旧版环绕图片结构。' }),
        M3009: Object.freeze({ severity: 'info', stage: 'plan', message: '图片已经处理。', cause: '节点存在幂等标记属性。', action: '不重复包裹或修改。' }),

        D4001: Object.freeze({ severity: 'error', stage: 'apply', message: '图片节点没有父节点。', cause: '无法创建包装节点，图片可能已被移除。', action: '跳过该图片。' }),
        D4002: Object.freeze({ severity: 'error', stage: 'apply', message: '缺少 ownerDocument。', cause: '节点不属于可创建元素的 Document。', action: '跳过该图片。' }),
        D4003: Object.freeze({ severity: 'error', stage: 'apply', message: '定位目标无效。', cause: '策略未能产生可设置样式的 Element。', action: '跳过该图片。' }),
        D4004: Object.freeze({ severity: 'error', stage: 'apply', message: '检测到 DOM 循环移动风险。', cause: '目标容器位于待移动节点内部。', action: '禁止移动并跳过。' }),
        D4005: Object.freeze({ severity: 'error', stage: 'apply', message: '应用 VML 布局失败。', cause: 'DOM 包裹、样式写入、节点移动或属性写入抛出异常。', action: '保留可恢复的现有结构并记录原始异常。' }),
        D4006: Object.freeze({ severity: 'info', stage: 'apply', message: '占位 table 未删除。', cause: 'table 仍包含可见或业务节点，不满足安全删除条件。', action: '保留 table。' }),

        S5001: Object.freeze({ severity: 'warning', stage: 'source', message: '普通图片 src 无效。', cause: 'src 为空或无法提取文件名。', action: '不使用该图片更新 VML 源路径。' }),
        S5002: Object.freeze({ severity: 'warning', stage: 'source', message: 'VML 图片源无法替换。', cause: 'shape ID、imagedata 或目标文件名缺失。', action: '保留原 VML src。' }),
        S5003: Object.freeze({ severity: 'error', stage: 'source', message: 'VML 图片源改写失败。', cause: '读取或重写 innerHTML 时发生异常。', action: '终止源路径改写并记录原始异常。' }),
        S5004: Object.freeze({ severity: 'warning', stage: 'source', message: '同一 shape ID 对应多个普通图片源。', cause: '多个 img 使用相同 v:shapes ID，但文件名不同。', action: '保留首次出现的图片源。' }),

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

    function normalizeOptions(options, report) {
        if (options == null) return { ...DEFAULT_OPTIONS };
        if (typeof options !== 'object' || Array.isArray(options)) {
            addIssue(report, 'I1004', { receivedType: Array.isArray(options) ? 'array' : typeof options });
            return { ...DEFAULT_OPTIONS };
        }

        const normalized = { ...DEFAULT_OPTIONS };
        const booleanKeys = [
            'rewriteImageSources', 'convertVmlParagraphs', 'processAllWrapTypes',
            'removeEmptyPlaceholderTables', 'skipDirectDivImages', 'debug'
        ];
        booleanKeys.forEach(key => {
            if (!(key in options)) return;
            if (typeof options[key] === 'boolean') normalized[key] = options[key];
            else addIssue(report, 'I1005', { option: key, receivedType: typeof options[key] });
        });

        if ('markAttribute' in options) {
            if (typeof options.markAttribute === 'string' && /^data-[a-z_][a-z0-9_.:-]*$/i.test(options.markAttribute)) {
                normalized.markAttribute = options.markAttribute;
            } else {
                addIssue(report, 'I1006', { option: 'markAttribute', receivedValue: String(options.markAttribute) });
            }
        }

        if ('maxCoordinatePt' in options) {
            if (Number.isFinite(options.maxCoordinatePt) && options.maxCoordinatePt > 0) {
                normalized.maxCoordinatePt = options.maxCoordinatePt;
            } else {
                addIssue(report, 'I1007', { option: 'maxCoordinatePt', receivedValue: options.maxCoordinatePt });
            }
        }
        return normalized;
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

    function isShapeEligible(shape, options) {
        if (!shape.hasWrap || options.processAllWrapTypes) return true;
        // 保持旧实现的筛选语义：浮于文字上/衬于文字下相关锚点进入定位流程。
        return shape.anchor.x === 'margin' || shape.anchor.y === 'page';
    }

    function isExcludedContainer(element) {
        return element.classList.contains('WordSection1') ||
            element.id === 'mail-content' ||
            element.id === 'mail-header';
    }

    function isNearestVmlContainer(element) {
        // 只在更深层的“候选定位容器”也含 VML 时跳过当前元素。
        // span/table 等非候选子节点即使承载条件注释，也不应导致 shape 无处解析。
        return !Array.from(element.querySelectorAll('div, li, h2')).some(descendant =>
            containsVmlCondition(descendant.innerHTML)
        );
    }

    function extractShapes(root, options, report) {
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
                if (isShapeEligible(shape, options)) records.push(shape);
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
        const right = Math.max(...valid.map(shape => shape.boxPt.left + shape.boxPt.width));
        const bottom = Math.max(...valid.map(shape => shape.boxPt.top + shape.boxPt.height));

        return {
            left,
            top,
            width: Math.max(0, right - left),
            height: Math.max(0, bottom - top),
            // 沿用旧代码的最小层级策略，后续可用真实多 shape 样本调整。
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

    function isLayoutWithinBounds(layout, maxCoordinatePt) {
        if (!layout) return false;
        const values = [layout.left, layout.top, layout.width, layout.height, layout.zIndex];
        return values.every(Number.isFinite) &&
            layout.width > 0 && layout.height > 0 &&
            [layout.left, layout.top, layout.width, layout.height].every(value =>
                Math.abs(value) <= maxCoordinatePt
            );
    }

    function createImagePlan(root, image, shapeById, options, report) {
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
        if (!isLayoutWithinBounds(layout, options.maxCoordinatePt)) {
            addIssue(report, 'M3007', { shapeIds: ids, layout, maxCoordinatePt: options.maxCoordinatePt });
            return { image, action: 'skip', reason: 'M3007' };
        }

        if (options.skipDirectDivImages && image.parentElement?.tagName.toLowerCase() === 'div') {
            addIssue(report, 'M3008', { shapeIds: ids });
            return { image, action: 'skip', reason: 'M3008' };
        }

        let target = null;
        let strategy = 'new-wrapper';
        const directParent = image.parentElement;

        if (directParent?.tagName.toLowerCase() === 'span') {
            target = directParent;
            strategy = 'existing-span';
        } else if (directParent?.tagName.toLowerCase() === 'a') {
            target = findAncestorWithin(directParent.parentElement, 'span', container);
            if (target) {
                strategy = 'link-existing-span';
            } else {
                target = directParent;
                strategy = 'link-wrapper';
            }
        }

        return {
            image,
            ids,
            shapes,
            container,
            layout,
            target,
            strategy,
            action: 'transform'
        };
    }

    function buildPlans(root, shapeById, options, report) {
        const images = Array.from(root.querySelectorAll('img')).filter(image =>
            image.hasAttribute('v:shapes')
        );
        report.stats.images = images.length;

        return images.map(image => {
            if (image.hasAttribute(options.markAttribute)) {
                addIssue(report, 'M3009', { shapeIds: splitShapeIds(image.getAttribute('v:shapes')) });
                return { image, action: 'already-processed', reason: 'M3009' };
            }
            return createImagePlan(root, image, shapeById, options, report);
        }).filter(Boolean);
    }

    function ensurePositioningContainer(container) {
        const position = container.style.getPropertyValue('position').toLowerCase();
        if (!position || position === 'static') {
            container.style.setProperty('position', 'relative');
        }
    }

    function setLayoutStyles(element, layout) {
        element.style.setProperty('position', 'absolute');
        element.style.setProperty('left', `${layout.left}pt`);
        element.style.setProperty('top', `${layout.top}pt`);
        element.style.setProperty('width', `${layout.width}pt`);
        element.style.setProperty('height', `${layout.height}pt`);
        element.style.setProperty('z-index', String(layout.zIndex));
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

    function applyPlan(plan, options, report) {
        if (plan.action !== 'transform') {
            report.stats.skipped += 1;
            return;
        }

        const { image, container, layout } = plan;
        const originalTable = typeof image.closest === 'function' ? image.closest('table') : null;

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
            } else if (plan.strategy === 'link-wrapper') {
                positionedNode = wrapElement(plan.target);
            }

            if (!positionedNode?.style || typeof positionedNode.setAttribute !== 'function') {
                throw createCodedError('D4003');
            }
            if (positionedNode.contains(container)) throw createCodedError('D4004');

            // 只覆盖本模块负责的布局属性，保留字体、颜色等邮件原始样式。
            setLayoutStyles(positionedNode, layout);
            positionedNode.setAttribute(options.markAttribute, '1');
            positionedNode.setAttribute('data-vml-shape-ids', plan.ids.join(' '));
            image.setAttribute(options.markAttribute, '1');
            image.classList.add('vml-controlled');

            // 绝对定位节点必须位于其定位容器内。对已在容器内的节点不做搬移。
            if (!container.contains(positionedNode)) {
                container.appendChild(positionedNode);
            }

            if (options.removeEmptyPlaceholderTables && originalTable && originalTable.isConnected) {
                // 定位容器或定位节点仍在原 table 内时绝不能删除，否则会连同有效内容一起移除。
                const containsLiveLayout = originalTable.contains(container) || originalTable.contains(positionedNode);
                if (!containsLiveLayout && isEmptyPlaceholderTable(originalTable, positionedNode)) {
                    originalTable.remove();
                } else {
                    addIssue(report, 'D4006', {
                        shapeIds: plan.ids,
                        reason: containsLiveLayout ? 'table-contains-live-layout' : 'table-not-empty'
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
            addIssue(report, 'X9001', { entry: 'changeVMLSrcs', originalError: error?.message || String(error) });
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
    function parseVMLImages(input, suppliedOptions, suppliedReport) {
        const report = resolveReport(suppliedReport);
        try {
            const options = normalizeOptions(suppliedOptions, report);
            const root = validateRoot(input, report);
            if (!root || !root.innerHTML) return report;

            const shapes = extractShapes(root, options, report);
            const shapeById = buildShapeIndex(shapes, report);
            const plans = buildPlans(root, shapeById, options, report);
            report.stats.matched = plans.filter(plan => plan.action === 'transform').length;
            plans.forEach(plan => applyPlan(plan, options, report));

            if (options.debug && report.issues.length) {
                global.console?.warn?.('[OutlookVMLParser]', report);
            }
            return report;
        } catch (error) {
            addIssue(report, 'X9001', { entry: 'parseVMLImages', originalError: error?.message || String(error) });
            return report;
        }
    }

    /**
     * 可选组合入口，适用于调用方明确需要在同一 DOM 上连续执行多个阶段的场景。
     * 常规业务仍应分别调用 changeVMLSrcs，或 VML_P2DIV + parseVMLImages。
     */
    function transform(input, suppliedOptions) {
        const report = createReport();
        try {
            const options = normalizeOptions(suppliedOptions, report);
            const root = validateRoot(input, report);
            if (!root || !root.innerHTML) return report;

            if (options.convertVmlParagraphs) {
                root.innerHTML = convertVmlParagraphs(root.innerHTML, report);
            }
            if (options.rewriteImageSources) {
                rewriteVmlImageSources(root, report);
            }
            return parseVMLImages(root, options, report);
        } catch (error) {
            addIssue(report, 'X9001', { entry: 'transform', originalError: error?.message || String(error) });
            return report;
        }
    }

    const api = Object.freeze({
        VERSION,
        ERROR_CODES,
        createReport,
        transform,
        changeVMLSrcs,
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

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordAddAIFormula_Final.Services;
using WordAddAIFormula_Final.Infrastructure.Converters;
using WordAddAIFormula_Final.Interfaces;

namespace WordAIFormula_Final
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // 1. 获取剪贴板
            if (!Clipboard.ContainsText())
            {
                MessageBox.Show("剪贴板为空！");
                return;
            }
            string rawText = Clipboard.GetText();

            // 2. 准备工作
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app.Selection;
            IFormulaConverter converter = new MathTypeConverter(app);

            // ================== 核心升级：智能正则解析 ==================
            // 这个正则会匹配：
            // 1. $$...$$ (换行公式)
            // 2. \[...\] (换行公式)
            // 3. \(...\) (行内公式)
            // 4. $...$   (行内公式，排除转义的 \$)
            string pattern = @"(\$\$.*?\$\$|\\\[.*?\\\]|\\\(.*?\\\)|(?<!\\)\$(?:\\.|[^$])*?(?<!\\)\$)";

            // 使用 Regex.Split 把文本切分成：[文字, 公式, 文字, 公式...]
            // RegexOptions.Singleline 让点号(.)也能匹配换行符
            string[] segments = Regex.Split(rawText, pattern, RegexOptions.Singleline);

            try
            {
                // 关闭屏幕更新，防止处理过程中屏幕疯狂闪烁
                app.ScreenUpdating = false;

                foreach (string segment in segments)
                {
                    if (string.IsNullOrEmpty(segment)) continue;

                    // 检查这一段是不是公式（看是否以特定符号开头和结尾）
                    bool isFormula = LatexCleaner.ContainsLatexDelimiters(segment);

                    if (isFormula)
                    {
                        // === 是公式：清洗 -> 插入 -> 选中 -> 转换 ===

                        // 1. 清洗 (去掉 $$, \[, \] 等外壳，拿到纯核心)
                        string coreLatex = LatexCleaner.CleanLatex(segment);

                        // 2. 重新包装为 MathType 喜欢的标准格式 \[ ... \]
                        // 技巧：MathType 对 \[...\] 识别率最高
                        string mathTypeInput = "\\[" + coreLatex + "\\]";

                        // 3. 记录当前插入点，插入文本
                        Word.Range range = selection.Range;
                        range.Text = mathTypeInput;

                        // 4. 选中刚才插入的这段 LaTeX 代码
                        // 注意：这里需要精确选中，否则可能把旁边的文字也卷进去
                        range.Select(); // 这一步至关重要

                        // 5. 呼叫 MathType 转换
                        converter.ConvertToFormat(coreLatex);

                        // 6. 转换完后，光标通常会自动在公式后面，我们需要把选区折叠到末尾，以便继续插入下一段
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                    else
                    {
                        // === 是普通文字：直接插入 ===
                        selection.TypeText(segment);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理混合文本时出错: " + ex.Message);
            }
            finally
            {
                // 恢复屏幕更新
                app.ScreenUpdating = true;
            }
        }
    }
}
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FormSolution
{

    public partial class Ribbon1
    {
        private int InputLength = 38;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            CheckIfCellWraps();
        }

        private void CheckIfCellWraps()
        {
            Word.Application app = Globals.ThisAddIn.Application;

            try
            {
                Word.Selection selection = app.Selection;

                // 确保当前选中的是表格单元格
                if (selection.Information[Word.WdInformation.wdWithInTable])
                {
                    var currentCell = selection.Cells[1];

                    Form1 inputForm = new Form1();

                    if (inputForm.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }
                    // 清空单元格内容
                    currentCell.Range.Text = string.Empty;
                    var text = inputForm.InputText.Replace("\r", "");

                    ProgressForm progressForm = new ProgressForm();

                    // 使任务完成后关闭等待框
                    Task backgroundTask = Task.Run(() =>
                    {
                        foreach (char c in text)
                        {
                            if (c == '\n')
                            {
                                currentCell = MoveToNextCellOrAddRow(currentCell);
                                if (currentCell == null)
                                {
                                    break;
                                }
                                continue;
                            }

                            bool heightChanged = HasCursorHeightChangedAfterTyping(c.ToString());

                            if (heightChanged)
                            {
                                currentCell = MoveToNextCellOrAddRow(currentCell);
                                if (currentCell == null)
                                {
                                    break;
                                }
                            }
                        }
                    });

                    // 显示等待框，并等待任务完成
                    progressForm.Shown += async (sender, args) =>
                    {
                        // 等待任务完成后关闭窗口
                        await backgroundTask;
                        progressForm.Close();
                    };

                    progressForm.ShowDialog();
                }
                else
                {
                    MessageBox.Show("请先选择一个表格单元格");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 移动到当前表格的下一列，如果没有下一列，则向下新建一行并移动到下一行的第一个单元格。
        /// </summary>
        /// <param name="currentCell">当前单元格。</param>
        /// <returns>移动后的单元格；如果失败，则返回 null。</returns>
        private static Word.Cell MoveToNextCellOrAddRow(Word.Cell currentCell)
        {
            try
            {
                Word.Cell newCell;
                // 获取当前单元格的表格和行
                Word.Table table = currentCell.Range.Tables[1];
                Word.Row currentRow = currentCell.Row;
                // 如果当前单元格有下一行，直接返回下一行
                if (currentCell.Next != null)
                {
                    Word.Cell nextCell = currentRow.Cells[currentRow.Cells.Count].Next;
                    // 检查下一行是否有内容
                    if (string.IsNullOrEmpty(nextCell.Range.Text.Trim().Replace("\a", "")))
                    {
                        currentCell.Next.Select();
                        return currentCell.Next;
                    } else
                    {
                        // 在当前行下方添加新行
                        Word.Row newRow = table.Rows.Add(currentRow.Next);

                        // 获取新行的第一个单元格
                        newCell = newRow.Cells[1];

                        // 关闭粗体
                        newCell.Range.Font.Bold = 0;

                        // 取消编号
                        newCell.Range.ListFormat.RemoveNumbers();

                        // 移动光标到新行的第一个单元格
                        newCell.Select();
                        return newCell;
                    }
                }

                // 检查是否为最后一行，若是则新增一行
                if (currentRow.IsLast)
                {
                    table.Rows.Add(); // 在表格末尾新增一行
                }

                // 移动到新增行的第一个单元格
                newCell = table.Rows[table.Rows.Count].Cells[1];
                newCell.Select();
                return newCell;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法移动到下一单元格或新增行: {ex.Message}");
                return null;
            }
        }




        /// <summary>
        /// 检查键入指定文本后光标的垂直高度是否发生变化。
        /// </summary>
        /// <param name="inputText">要插入的文本内容。</param>
        /// <returns>如果光标高度发生变化，则返回 true；否则返回 false。</returns>
        private static bool HasCursorHeightChangedAfterTyping(string inputText)
        {
            try
            {
                // 获取当前 Word 应用程序实例和选区
                Word.Application application = Globals.ThisAddIn.Application;
                Word.Selection selection = application.Selection;

                // 记录插入前的光标起始位置
                int initialPosition = selection.Range.Start;

                // 获取光标初始的垂直位置
                float initialVerticalPosition = selection.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];

                // 插入指定文本
                selection.TypeText(inputText);

                // 获取光标插入文本后的垂直位置
                float newVerticalPosition = selection.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];

                // 判断高度是否发生变化（允许小范围误差）
                bool hasChanged = Math.Abs(initialVerticalPosition - newVerticalPosition) > 0.01f;

                // 如果高度发生变化，则删除插入的文本
                if (hasChanged)
                {
                    Word.Range insertedRange = selection.Document.Range(initialPosition, initialPosition + inputText.Length);
                    insertedRange.Delete(); // 删除刚刚插入的文本
                }
                return hasChanged;
            }
            catch (Exception ex)
            {
                // 异常处理并返回 false
                MessageBox.Show($"发生错误: {ex.Message}");
                return false;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            AutoLengthInput();
        }

        private void AutoLengthInput()
        {
            Word.Application app = Globals.ThisAddIn.Application;

            try
            {
                Word.Selection selection = app.Selection;

                // 确保当前选中的是表格单元格
                if (selection.Information[Word.WdInformation.wdWithInTable])
                {
                    var currentCell = selection.Cells[1];

                    Form1 inputForm = new Form1();

                    if (inputForm.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }
                    // 清空单元格内容
                    currentCell.Range.Text = string.Empty;
                    var text = inputForm.InputText.Replace("\r", "");

                    var lines = text.Split('\n');

                    ProgressForm progressForm = new ProgressForm();

                    // 使任务完成后关闭等待框
                    Task backgroundTask = Task.Run(() =>
                    {
                        foreach (string line in lines)
                        {
                            int charIndex = 0;
                            Debug.WriteLine(line);
                            List<char> charList = new List<char>(line);
                            while (charIndex < charList.Count - 1)
                            {
                                // 判断是否换行
                                bool isLineFeed = false;
                                // 10 -> 9 -> 8，8长度的字符，触发表格单元格不换行，则退出循环
                                while (HasCursorHeightChangedAfterTyping(string.Join("", charList.GetRange(charIndex, Math.Min(InputLength, charList.Count - charIndex)))))
                                {
                                    // 如果换行
                                    isLineFeed = true;
                                    // 添加字符宽度
                                    InputLength--;
                                }
                                // 表格输入成功
                                // 将起始字符调整为，已读取字符之后
                                charIndex += InputLength + 1;
                                if (charIndex >= charList.Count)
                                {
                                    break;
                                }
                                if (!isLineFeed)
                                {
                                    // 如果表格未换行，逐个输入，直到换行后结束
                                    // 每次输入1个字符，1 -> 2 -> 3，直到触发表格单元格换行，退出循环
                                    while (!HasCursorHeightChangedAfterTyping(string.Join("", charList.GetRange(charIndex, 1))))
                                    {
                                        // 如果没换行
                                        // 字符被输入，索引+1
                                        charIndex++;
                                        InputLength++;
                                        if (charIndex >= charList.Count)
                                        {
                                            break;
                                        }
                                    }
                                    // 退出循环，字符换行，因换行字符被删除，所以输入宽度-1
                                    InputLength--;
                                    // 表格换行
                                    currentCell = MoveToNextCellOrAddRow(currentCell);
                                    //if (charIndex + InputLength > charList.Count)
                                    //{
                                    //    InputLength = charList.Count - charIndex;
                                    //}
                                }
                            }
                            currentCell = MoveToNextCellOrAddRow(currentCell);
                        }
                    });

                    // 显示等待框，并等待任务完成
                    progressForm.Shown += async (sender, args) =>
                    {
                        // 等待任务完成后关闭窗口
                        await backgroundTask;
                        progressForm.Close();
                    };

                    progressForm.ShowDialog();
                }
                else
                {
                    MessageBox.Show("请先选择一个表格单元格");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}");
            }
        }
    }
}

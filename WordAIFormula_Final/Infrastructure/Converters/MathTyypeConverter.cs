using System;
using Word = Microsoft.Office.Interop.Word;
using WordAddAIFormula_Final.Interfaces;

namespace WordAddAIFormula_Final.Infrastructure.Converters
{
    public class MathTypeConverter : IFormulaConverter
    {
        private readonly Word.Application _wordApplication;

        public MathTypeConverter(Word.Application wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
        }

        public bool ConvertToFormat(string latex)
        {
            if (string.IsNullOrWhiteSpace(latex)) return false;

            string[] commands = new string[]
            {
                "MTCommand_TexToggle",
                "MathType Commands 2016.dotm!MTCommand_TexToggle",
                "MathType Commands 6 For Word.dotm!MTCommand_TexToggle",
                "MathType Commands.dotm!MTCommand_TexToggle"
            };

            foreach (var cmd in commands)
            {
                try
                {
                    _wordApplication.Run(cmd);
                    return true;
                }
                catch
                {
                    continue;
                }
            }

            System.Windows.Forms.MessageBox.Show(
                "Failed to convert with MathType command.\nPlease try manually clicking MathType -> 'Toggle TeX' at the top of Word to confirm if the software is functioning properly.",
                "Conversion Failed",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);

            return false;
        }
    }
}
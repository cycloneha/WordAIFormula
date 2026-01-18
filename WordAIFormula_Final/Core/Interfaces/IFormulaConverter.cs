using System;

namespace WordAddAIFormula_Final.Interfaces
{
    public interface IFormulaConverter
    {
        bool ConvertToFormat(string latex);
    }
}
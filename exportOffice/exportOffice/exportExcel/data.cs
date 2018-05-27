using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exportexcel
{
    class data
    {
        //序号	引导词	要素		偏离		可能的原因		后果		安全措施		注释		建议措施			责任人
        private int id;
        private string guideword;
        private string key;
        private string deviate;
        private string possiblecause;
        private string consequence;
        private string safetymeasures;
        private string annotation;
        private string suggestionmeasure;
        private string responsibilityperson;

        public data(int id, string guideword, string key, string deviate, string possiblecause, string consequence, string safetymeasures, string annotation, string suggestionmeasure, string responsibilityperson)
        {
            this.Id = id;
            this.Guideword = guideword;
            this.Key = key;
            this.Deviate = deviate;
            this.Possiblecause = possiblecause;
            this.Consequence = consequence;
            this.Safetymeasures = safetymeasures;
            this.Annotation = annotation;
            this.Suggestionmeasure = suggestionmeasure;
            this.Responsibilityperson = responsibilityperson;
        }

        public int Id { get => id; set => id = value; }
        public string Guideword { get => guideword; set => guideword = value; }
        public string Key { get => key; set => key = value; }
        public string Deviate { get => deviate; set => deviate = value; }
        public string Possiblecause { get => possiblecause; set => possiblecause = value; }
        public string Consequence { get => consequence; set => consequence = value; }
        public string Safetymeasures { get => safetymeasures; set => safetymeasures = value; }
        public string Annotation { get => annotation; set => annotation = value; }
        public string Suggestionmeasure { get => suggestionmeasure; set => suggestionmeasure = value; }
        public string Responsibilityperson { get => responsibilityperson; set => responsibilityperson = value; }
    }
}

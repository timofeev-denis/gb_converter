using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GBConverter {
    struct Appeal {
        public string id;
        public string numb;
        public string f_date;
        public string content;
        public string confirmation;
        public string subjcode;
        public string parent_id;
        public string measures;
        public string party;
        public string declarant_type;
        public string theme;
        public string created;
        public string content_cik;
        public string executor_id;
        public bool isParent;
        public ArrayList multi;
        public void init() {
            id = null;
            numb = null;
            f_date = null;
            content = null;
            confirmation = null;
            subjcode = null;
            parent_id = null;
            measures = null;
            party = null;
            declarant_type = null;
            theme = null;
            created = null;
            content_cik = null;
            executor_id = null;
            isParent = false;
            multi = new ArrayList();
        }
        //public void Prepare(string id, string parent_id, string subjcode, string numb, string f_date, ArrayList appealMulti, ArrayList ik_subjcodes) {
        public void Prepare(string id, string parent_id, string subjcode, string numb, string f_date) {
            this.id = id;
            this.parent_id = parent_id;
            this.subjcode = subjcode;
            this.numb = numb;
            this.f_date = f_date;
            /*
            this.multi = new ArrayList();
            foreach (string[] str in appealMulti) {
                this.multi.Add(new string[] { str[0], str[1], str[2] });
            }
            int SubjCodeIndex = 0;
            foreach (string sc in ik_subjcodes) {
                this.multi.Add(new string[] { "ik_subjcode", sc, SubjCodeIndex.ToString() });
                SubjCodeIndex++;
            }
            */
        }
    }
}

using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace GBConverter {
    class Declarant {
        private string id = "";
        private string f_name;
        private string m_name;
        private string l_name;
        private string type;
        private string party;
        private string info;
        /// <summary>
        /// <para>Заявители равны, если совпадают все их поля кроме id.</para>
        /// </summary>
        /// <param name="obj">Экземпляр класса Declarant</param>
        /// <returns></returns>
        public override bool Equals(object obj) {
            Declarant declarantObj = obj as Declarant;
            if (declarantObj == null) {
                return false;
            }
            
            if (f_name.Equals(declarantObj.f_name) &&
                m_name.Equals(declarantObj.m_name) &&
                l_name.Equals(declarantObj.l_name) &&
                type.Equals(declarantObj.type) &&
                party.Equals(declarantObj.party) &&
                info.Equals(declarantObj.info)) {
                    return true;
            }
            return false;
        }
        public override int GetHashCode() {
            return Tuple.Create(f_name, m_name, l_name, type, party, info).GetHashCode();
        }
        public Declarant(string firstName, string middleName, string lastName, string type, string party, string info, string id = "") {
            this.f_name = firstName;
            this.m_name = middleName;
            this.l_name = lastName;
            this.type = type;
            this.party = party;
            this.info = info;
            this.id = id;
        }
        public Declarant(string type, string party, string info) {
            this.type = type;
            this.party = party;
            this.info = info;
        }
        public void SetFIO(string fio) {
            this.f_name = "-";
            this.m_name = "-";
            this.l_name = "-";

            string FullPattern = @"\p{IsCyrillic}\.\p{IsCyrillic}\. \w+";

            if (fio.ToLower() == "коллективное обращение") {
                this.f_name = "коллективное";
                this.m_name = "";
                this.l_name = "обращение";
            } else if (Regex.IsMatch(fio, FullPattern)) {
                this.f_name = fio.Substring(0, 1);
                this.m_name = fio.Substring(2, 1);
                this.l_name = fio.Substring(5);
            } else {
                this.f_name = fio.Substring(0, 1);
                this.m_name = "";
                this.l_name = fio.Substring(3);
            }
        }
        public string GetFIO() {
            return this.f_name + "." + ((this.m_name == "") ? "" : this.m_name + ".") + " " + this.l_name;
        }
        public string GetID() {
            return this.id;
        }
        public bool SetID(string newID) {
            if (this.id == "" || this.id.ToString().Contains("fake")) {
                this.id = newID;
                return true;
            } else {
                return false;
            }
        }
        /// <summary>
        /// <para>Сохранение заявителя в БД</para>
        /// <para>Возвращает id нового заявителя</para>
        /// </summary>
        /// <returns></returns>
        public string SaveToDB(DateTime ConvertDate) {
            OracleConnection conn = DB.GetConnection();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            OracleDataReader dr = null;
            string NewDeclarantID = "";
            //

            cmd.CommandText = "select CONCAT('100', LPAD(AKRIKO.seq_cat_declarants.NEXTVAL,7,'0')) from dual";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                cmd.Dispose();
                throw;
            }
            dr.Read();
            if (dr.IsDBNull(0)) {
                cmd.Dispose();
                throw new Exception("Не удалось сформировать id заявителя.");
            } else {
                NewDeclarantID = dr.GetString(0);
                this.id = NewDeclarantID;
            }
            
            //cmd.CommandType = CommandType.Text;
            cmd.CommandText = "INSERT INTO akriko.cat_declarants (id,l_name,f_name,m_name,type,party,info,created) " +
                "VALUES(:newdeclarantid, :lname, :fname, :mname, :type, :party, :info, :created)";

            //OracleCommand command = new OracleCommand(query, conn);
            cmd.Parameters.Add(":newdeclarantid", NewDeclarantID);
            cmd.Parameters.Add(":lname", this.l_name);
            cmd.Parameters.Add(":fname", this.f_name);
            cmd.Parameters.Add(":mname", this.m_name);
            cmd.Parameters.Add(":type", this.type);
            cmd.Parameters.Add(":party", this.party);
            cmd.Parameters.Add(":info", this.info);
            cmd.Parameters.Add(":created", ConvertDate);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            return NewDeclarantID;
        }
    }
}

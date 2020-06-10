using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Tests
{
    [TestFixture] 
    public class SelectStatementTests
    {
        [Test]
        public void SelectStatement_Fields_AddFieldsByNames()
        {
            var s = new SelectStatement();
            s.Fields.Add("F1", "F2", "F3");

            // TODO: Was bringt dieser Test?
            // ... Das was oben steht: er prüft, ob Add(...) aus der Fields-Auflistung funktioniert.
            // Soll man so einen Test später entfernen, wenn die Programmierung so einer Klasse abgeschlossen ist? 
            // ... ich bin der Meinung, dass man das nicht machen soll.
            var f = (FieldsStatement) s;
            Assert.AreEqual(s.Fields.Count, f.Fields.Count);
        }
    }
}

[System.Reflection.Assembly]::LoadWithPartialName("System.web")

##### CSHARP CLASS FOR DATATABLE HANDLING #####
$Assem = (
  "System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "System.Dynamic, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
  "System.Management.Automation, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35",
  "System.Runtime.Serialization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35",
  "System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "Microsoft.CSharp, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
)

$Source = @" 
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Web.Script.Serialization;
using System.Runtime.Serialization;
using System.Xml;

namespace Fluxx.GrantMaker { 
    public static class DataTableUtility {   

        public static string SerialzeToXml(dynamic records) {
            StringWriter sw = null;

            if (records.Length == 0) {
                return null;
            }

            // determine record type from first non-null item in the array
            Type recordType = null;
            for (int i=0; i < records.Length; i++) { 
                try {
                    recordType = records[i].GetType();
                    break;
                } catch {
                }
            }

            sw = new StringWriter();
            using (XmlTextWriter writer = new XmlTextWriter(sw)) {
                writer.WriteStartElement("Data");
                if (recordType == typeof(System.String)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteString(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.Boolean)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.DateTime)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.Decimal)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.Double)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.Int32)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.Int64)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else if (recordType == typeof(System.Single)) {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        writer.WriteValue(record);
                        writer.WriteEndElement();
                    }
                } else {
                    foreach (var record in records) {
                        writer.WriteStartElement("Object");
                        foreach (var attribute in record) {
                            if (attribute.Value != null) {
                                writer.WriteAttributeString(attribute.Key,Convert.ToString(attribute.Value));
                            }
                        }
                        writer.WriteEndElement();
                    }
                }
                writer.WriteEndElement();
                writer.Flush();
            }
            return sw.ToString();
        }

        public static void ProcessDates(string dateCols, ref DataTable dataTable) {

            string[] columns = dateCols.Split(new string[] { "," }, StringSplitOptions.None);
            DataTable newDataTable = new DataTable();

            foreach(DataColumn column in dataTable.Columns) {

                if (Array.IndexOf(columns, column.ColumnName) >= 0) {
                    var col = new DataColumn(column.ColumnName, typeof(DateTime));
                    col.DateTimeMode = DataSetDateTime.Utc;
                    newDataTable.Columns.Add(col);
                    var logString = DateTime.Now.ToString("yyyyMMdd.HHmmss");
                    Console.WriteLine(logString + " Fluxx.GrantMaker.DataTableUtility - Formatting " + column.ColumnName + " as DateTime");
                } else {
                    newDataTable.Columns.Add(column.ColumnName, column.DataType);
                }
            }

            newDataTable.Load(dataTable.CreateDataReader(), System.Data.LoadOption.OverwriteChanges);
            dataTable = newDataTable;
        }

        public static void ProcessRecords(String json, ref DataTable dataTable) {
            // loading JSON
            JavaScriptSerializer ser = new JavaScriptSerializer();
            ser.MaxJsonLength = Int32.MaxValue;
            dynamic records = ser.Deserialize<Object>(json);

            if (records == null) {
                throw new Exception(json);
                throw new Exception("record data not found");
            }

            var logString = DateTime.Now.ToString("yyyyMMdd.HHmmss");
            Console.WriteLine(logString + " Fluxx.GrantMaker.DataTableUtility - " + records.Length.ToString() + " records loaded");

            var c = 0;
            foreach (var record in records) {
                c += 1;
                var columnValues = new Dictionary<String,Object>();

                foreach (var key in record.Keys) {

                    Type new_col_type = null;
                    String new_col_name = key;
                    dynamic recordValue = record[key];
                    if(recordValue != null) {

                        new_col_type = recordValue.GetType();

                        if (new_col_type == typeof(System.Collections.Generic.Dictionary<System.String,System.Object>)) {
                            recordValue = new Object[]{recordValue};
                            new_col_type = recordValue.GetType();
                        }

                        if (new_col_type == typeof(System.Object[])) {
                            // if no data skip the column
                            if (recordValue.Length == 0) {
                                continue;
                            }

                            new_col_name = key + "_id";
    
                            // determine record type from first non-null item in the array
                            Type recordType = null;
                            for (int i=0; i < recordValue.Length; i++) {
                                try {
                                    recordType = recordValue[i].GetType();
                                    break;
                                } catch {
                                }
                            }
                            if (recordType == null) {
                                logString = DateTime.Now.ToString("yyyyMMdd.HHmmss");
                                Console.WriteLine(logString + " Fluxx.GrantMaker.DataTableUtility - skipping - unable to determine record type (id=" + record["id"] + ",key=" + key + ")");
                                continue;
                            }
                            var testKey = false;
                            try { 
                                testKey = recordValue[0].ContainsKey("id"); 
                            } catch {
                                testKey = false;
                            }
                            if (recordType != typeof(System.String) && testKey) {
                                if (!dataTable.Columns.Contains(new_col_name)) {
                                    DataColumn new_col = new DataColumn();
                                    logString = DateTime.Now.ToString("yyyyMMdd.HHmmss");
                                    Console.WriteLine(logString + " Fluxx.GrantMaker.DataTableUtility - adding new column: " + new_col_name);
                                    new_col.DataType =  typeof(System.Int32);
                                    new_col.ColumnName = new_col_name;
                                    new_col.AllowDBNull = true;
                                    new_col.Unique = false;
                                    new_col.DefaultValue = DBNull.Value;
                                    dataTable.Columns.Add(new_col);
                                }
                                columnValues[new_col_name] = recordValue[0]["id"];
                            }
                            
                            new_col_name = key + "_xml";
                            if (!dataTable.Columns.Contains(new_col_name)) {
                                DataColumn new_col = new DataColumn();
                                logString = DateTime.Now.ToString("yyyyMMdd.HHmmss");
                                Console.WriteLine(logString + " Fluxx.GrantMaker.DataTableUtility - adding new column: " + new_col_name);
                                new_col.DataType =  typeof(System.String);
                                new_col.ColumnName = new_col_name;
                                new_col.AllowDBNull = true;
                                new_col.Unique = false;
                                new_col.DefaultValue = DBNull.Value;
                                dataTable.Columns.Add(new_col);
                            }
                            columnValues[new_col_name] = SerialzeToXml(recordValue);

                        } else {

                            if (!dataTable.Columns.Contains(new_col_name)) {
                                DataColumn new_col = new DataColumn();
                                new_col_name = key;
                                logString = DateTime.Now.ToString("yyyyMMdd.HHmmss");
                                Console.WriteLine(logString + " Fluxx.GrantMaker.DataTableUtility - adding new column: " + new_col_name);
                                new_col.DataType = new_col_type;
                                new_col.ColumnName = new_col_name;
                                new_col.AllowDBNull = true;
                                new_col.Unique = false;
                                new_col.DefaultValue = DBNull.Value;
                                dataTable.Columns.Add(new_col);
                            }
                            columnValues[new_col_name] = recordValue;

                        }
                    }
                }

                DataRow dataRow = dataTable.NewRow();

                foreach (var key in columnValues.Keys) {
                    if(columnValues[key] != null) {
                        dataRow[key] = columnValues[key];
                    }
                }

                dataTable.Rows.Add(dataRow);

            }
        }
    } 
} 
"@ 

##### HELPER FUNCTIONS #####

# Internal function used to log messages consistently.
function Log-Message {
param(
      [String]$Context
     ,[String]$Message
     )
   write-host([DateTime]::Now.ToString("yyyyMMdd.HHmmss") + " " + $Context + " - " + $Message)
}

# Internal function used to determine content types based on file extensions
function Get-MimeContentType([String]$extension){
	$result = ""
	switch ($extension) 
		{ 
                 "323" {$result = "text/h323"}
                 "aaf" {$result = "application/octet-stream"}
                 "aca" {$result = "application/octet-stream"}
                 "accdb" {$result = "application/msaccess"}
                 "accde" {$result = "application/msaccess"}
                 "accdt" {$result = "application/msaccess"}
                 "acx" {$result = "application/internet-property-stream"}
                 "afm" {$result = "application/octet-stream"}
                 "ai" {$result = "application/postscript"}
                 "aif" {$result = "audio/x-aiff"}
                 "aifc" {$result = "audio/aiff"}
                 "aiff" {$result = "audio/aiff"}
                 "application" {$result = "application/x-ms-application"}
                 "art" {$result = "image/x-jg"}
                 "asd" {$result = "application/octet-stream"}
                 "asf" {$result = "video/x-ms-asf"}
                 "asi" {$result = "application/octet-stream"}
                 "asm" {$result = "text/plain"}
                 "asr" {$result = "video/x-ms-asf"}
                 "asx" {$result = "video/x-ms-asf"}
                 "atom" {$result = "application/atom+xml"}
                 "au" {$result = "audio/basic"}
                 "avi" {$result = "video/x-msvideo"}
                 "axs" {$result = "application/olescript"}
                 "bas" {$result = "text/plain"}
                 "bcpio" {$result = "application/x-bcpio"}
                 "bin" {$result = "application/octet-stream"}
                 "bmp" {$result = "image/bmp"}
                 "c" {$result = "text/plain"}
                 "cab" {$result = "application/octet-stream"}
                 "calx" {$result = "application/vnd.ms-office.calx"}
                 "cat" {$result = "application/vnd.ms-pki.seccat"}
                 "cdf" {$result = "application/x-cdf"}
                 "chm" {$result = "application/octet-stream"}
                 "class" {$result = "application/x-java-applet"}
                 "clp" {$result = "application/x-msclip"}
                 "cmx" {$result = "image/x-cmx"}
                 "cnf" {$result = "text/plain"}
                 "cod" {$result = "image/cis-cod"}
                 "cpio" {$result = "application/x-cpio"}
                 "cpp" {$result = "text/plain"}
                 "crd" {$result = "application/x-mscardfile"}
                 "crl" {$result = "application/pkix-crl"}
                 "crt" {$result = "application/x-x509-ca-cert"}
                 "csh" {$result = "application/x-csh"}
                 "css" {$result = "text/css"}
                 "csv" {$result = "application/octet-stream"}
                 "cur" {$result = "application/octet-stream"}
                 "dcr" {$result = "application/x-director"}
                 "deploy" {$result = "application/octet-stream"}
                 "der" {$result = "application/x-x509-ca-cert"}
                 "dib" {$result = "image/bmp"}
                 "dir" {$result = "application/x-director"}
                 "disco" {$result = "text/xml"}
                 "dll" {$result = "application/x-msdownload"}
                 "dll.config" {$result = "text/xml"}
                 "dlm" {$result = "text/dlm"}
                 "doc" {$result = "application/msword"}
                 "docm" {$result = "application/vnd.ms-word.document.macroEnabled.12"}
                 "docx" {$result = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                 "dot" {$result = "application/msword"}
                 "dotm" {$result = "application/vnd.ms-word.template.macroEnabled.12"}
                 "dotx" {$result = "application/vnd.openxmlformats-officedocument.wordprocessingml.template"}
                 "dsp" {$result = "application/octet-stream"}
                 "dtd" {$result = "text/xml"}
                 "dvi" {$result = "application/x-dvi"}
                 "dwf" {$result = "drawing/x-dwf"}
                 "dwp" {$result = "application/octet-stream"}
                 "dxr" {$result = "application/x-director"}
                 "eml" {$result = "message/rfc822"}
                 "emz" {$result = "application/octet-stream"}
                 "eot" {$result = "application/octet-stream"}
                 "eps" {$result = "application/postscript"}
                 "etx" {$result = "text/x-setext"}
                 "evy" {$result = "application/envoy"}
                 "exe" {$result = "application/octet-stream"}
                 "exe.config" {$result = "text/xml"}
                 "fdf" {$result = "application/vnd.fdf"}
                 "fif" {$result = "application/fractals"}
                 "fla" {$result = "application/octet-stream"}
                 "flr" {$result = "x-world/x-vrml"}
                 "flv" {$result = "video/x-flv"}
                 "gif" {$result = "image/gif"}
                 "gtar" {$result = "application/x-gtar"}
                 "gz" {$result = "application/x-gzip"}
                 "h" {$result = "text/plain"}
                 "hdf" {$result = "application/x-hdf"}
                 "hdml" {$result = "text/x-hdml"}
                 "hhc" {$result = "application/x-oleobject"}
                 "hhk" {$result = "application/octet-stream"}
                 "hhp" {$result = "application/octet-stream"}
                 "hlp" {$result = "application/winhlp"}
                 "hqx" {$result = "application/mac-binhex40"}
                 "hta" {$result = "application/hta"}
                 "htc" {$result = "text/x-component"}
                 "htm" {$result = "text/html"}
                 "html" {$result = "text/html"}
                 "htt" {$result = "text/webviewhtml"}
                 "hxt" {$result = "text/html"}
                 "ico" {$result = "image/x-icon"}
                 "ics" {$result = "application/octet-stream"}
                 "ief" {$result = "image/ief"}
                 "iii" {$result = "application/x-iphone"}
                 "inf" {$result = "application/octet-stream"}
                 "ins" {$result = "application/x-internet-signup"}
                 "isp" {$result = "application/x-internet-signup"}
                 "IVF" {$result = "video/x-ivf"}
                 "jar" {$result = "application/java-archive"}
                 "java" {$result = "application/octet-stream"}
                 "jck" {$result = "application/liquidmotion"}
                 "jcz" {$result = "application/liquidmotion"}
                 "jfif" {$result = "image/pjpeg"}
                 "jpb" {$result = "application/octet-stream"}
                 "jpe" {$result = "image/jpeg"}
                 "jpeg" {$result = "image/jpeg"}
                 "jpg" {$result = "image/jpeg"}
                 "js" {$result = "application/x-javascript"}
                 "jsx" {$result = "text/jscript"}
                 "latex" {$result = "application/x-latex"}
                 "lit" {$result = "application/x-ms-reader"}
                 "lpk" {$result = "application/octet-stream"}
                 "lsf" {$result = "video/x-la-asf"}
                 "lsx" {$result = "video/x-la-asf"}
                 "lzh" {$result = "application/octet-stream"}
                 "m13" {$result = "application/x-msmediaview"}
                 "m14" {$result = "application/x-msmediaview"}
                 "m1v" {$result = "video/mpeg"}
                 "m3u" {$result = "audio/x-mpegurl"}
                 "man" {$result = "application/x-troff-man"}
                 "manifest" {$result = "application/x-ms-manifest"}
                 "map" {$result = "text/plain"}
                 "mdb" {$result = "application/x-msaccess"}
                 "mdp" {$result = "application/octet-stream"}
                 "me" {$result = "application/x-troff-me"}
                 "mht" {$result = "message/rfc822"}
                 "mhtml" {$result = "message/rfc822"}
                 "mid" {$result = "audio/mid"}
                 "midi" {$result = "audio/mid"}
                 "mix" {$result = "application/octet-stream"}
                 "mmf" {$result = "application/x-smaf"}
                 "mno" {$result = "text/xml"}
                 "mny" {$result = "application/x-msmoney"}
                 "mov" {$result = "video/quicktime"}
                 "movie" {$result = "video/x-sgi-movie"}
                 "mp2" {$result = "video/mpeg"}
                 "mp3" {$result = "audio/mpeg"}
                 "mpa" {$result = "video/mpeg"}
                 "mpe" {$result = "video/mpeg"}
                 "mpeg" {$result = "video/mpeg"}
                 "mpg" {$result = "video/mpeg"}
                 "mpp" {$result = "application/vnd.ms-project"}
                 "mpv2" {$result = "video/mpeg"}
                 "ms" {$result = "application/x-troff-ms"}
                 "msi" {$result = "application/octet-stream"}
                 "mso" {$result = "application/octet-stream"}
                 "mvb" {$result = "application/x-msmediaview"}
                 "mvc" {$result = "application/x-miva-compiled"}
                 "nc" {$result = "application/x-netcdf"}
                 "nsc" {$result = "video/x-ms-asf"}
                 "nws" {$result = "message/rfc822"}
                 "ocx" {$result = "application/octet-stream"}
                 "oda" {$result = "application/oda"}
                 "odc" {$result = "text/x-ms-odc"}
                 "ods" {$result = "application/oleobject"}
                 "one" {$result = "application/onenote"}
                 "onea" {$result = "application/onenote"}
                 "onetoc" {$result = "application/onenote"}
                 "onetoc2" {$result = "application/onenote"}
                 "onetmp" {$result = "application/onenote"}
                 "onepkg" {$result = "application/onenote"}
                 "p10" {$result = "application/pkcs10"}
                 "p12" {$result = "application/x-pkcs12"}
                 "p7b" {$result = "application/x-pkcs7-certificates"}
                 "p7c" {$result = "application/pkcs7-mime"}
                 "p7m" {$result = "application/pkcs7-mime"}
                 "p7r" {$result = "application/x-pkcs7-certreqresp"}
                 "p7s" {$result = "application/pkcs7-signature"}
                 "pbm" {$result = "image/x-portable-bitmap"}
                 "pcx" {$result = "application/octet-stream"}
                 "pcz" {$result = "application/octet-stream"}
                 "pdf" {$result = "application/pdf"}
                 "pfb" {$result = "application/octet-stream"}
                 "pfm" {$result = "application/octet-stream"}
                 "pfx" {$result = "application/x-pkcs12"}
                 "pgm" {$result = "image/x-portable-graymap"}
                 "pko" {$result = "application/vnd.ms-pki.pko"}
                 "pma" {$result = "application/x-perfmon"}
                 "pmc" {$result = "application/x-perfmon"}
                 "pml" {$result = "application/x-perfmon"}
                 "pmr" {$result = "application/x-perfmon"}
                 "pmw" {$result = "application/x-perfmon"}
                 "png" {$result = "image/png"}
                 "pnm" {$result = "image/x-portable-anymap"}
                 "pnz" {$result = "image/png"}
                 "pot" {$result = "application/vnd.ms-powerpoint"}
                 "potm" {$result = "application/vnd.ms-powerpoint.template.macroEnabled.12"}
                 "potx" {$result = "application/vnd.openxmlformats-officedocument.presentationml.template"}
                 "ppam" {$result = "application/vnd.ms-powerpoint.addin.macroEnabled.12"}
                 "ppm" {$result = "image/x-portable-pixmap"}
                 "pps" {$result = "application/vnd.ms-powerpoint"}
                 "ppsm" {$result = "application/vnd.ms-powerpoint.slideshow.macroEnabled.12"}
                 "ppsx" {$result = "application/vnd.openxmlformats-officedocument.presentationml.slideshow"}
                 "ppt" {$result = "application/vnd.ms-powerpoint"}
                 "pptm" {$result = "application/vnd.ms-powerpoint.presentation.macroEnabled.12"}
                 "pptx" {$result = "application/vnd.openxmlformats-officedocument.presentationml.presentation"}
                 "prf" {$result = "application/pics-rules"}
                 "prm" {$result = "application/octet-stream"}
                 "prx" {$result = "application/octet-stream"}
                 "ps" {$result = "application/postscript"}
                 "psd" {$result = "application/octet-stream"}
                 "psm" {$result = "application/octet-stream"}
                 "psp" {$result = "application/octet-stream"}
                 "pub" {$result = "application/x-mspublisher"}
                 "qt" {$result = "video/quicktime"}
                 "qtl" {$result = "application/x-quicktimeplayer"}
                 "qxd" {$result = "application/octet-stream"}
                 "ra" {$result = "audio/x-pn-realaudio"}
                 "ram" {$result = "audio/x-pn-realaudio"}
                 "rar" {$result = "application/octet-stream"}
                 "ras" {$result = "image/x-cmu-raster"}
                 "rf" {$result = "image/vnd.rn-realflash"}
                 "rgb" {$result = "image/x-rgb"}
                 "rm" {$result = "application/vnd.rn-realmedia"}
                 "rmi" {$result = "audio/mid"}
                 "roff" {$result = "application/x-troff"}
                 "rpm" {$result = "audio/x-pn-realaudio-plugin"}
                 "rtf" {$result = "application/rtf"}
                 "rtx" {$result = "text/richtext"}
                 "scd" {$result = "application/x-msschedule"}
                 "sct" {$result = "text/scriptlet"}
                 "sea" {$result = "application/octet-stream"}
                 "setpay" {$result = "application/set-payment-initiation"}
                 "setreg" {$result = "application/set-registration-initiation"}
                 "sgml" {$result = "text/sgml"}
                 "sh" {$result = "application/x-sh"}
                 "shar" {$result = "application/x-shar"}
                 "sit" {$result = "application/x-stuffit"}
                 "sldm" {$result = "application/vnd.ms-powerpoint.slide.macroEnabled.12"}
                 "sldx" {$result = "application/vnd.openxmlformats-officedocument.presentationml.slide"}
                 "smd" {$result = "audio/x-smd"}
                 "smi" {$result = "application/octet-stream"}
                 "smx" {$result = "audio/x-smd"}
                 "smz" {$result = "audio/x-smd"}
                 "snd" {$result = "audio/basic"}
                 "snp" {$result = "application/octet-stream"}
                 "spc" {$result = "application/x-pkcs7-certificates"}
                 "spl" {$result = "application/futuresplash"}
                 "src" {$result = "application/x-wais-source"}
                 "ssm" {$result = "application/streamingmedia"}
                 "sst" {$result = "application/vnd.ms-pki.certstore"}
                 "stl" {$result = "application/vnd.ms-pki.stl"}
                 "sv4cpio" {$result = "application/x-sv4cpio"}
                 "sv4crc" {$result = "application/x-sv4crc"}
                 "swf" {$result = "application/x-shockwave-flash"}
                 "t" {$result = "application/x-troff"}
                 "tar" {$result = "application/x-tar"}
                 "tcl" {$result = "application/x-tcl"}
                 "tex" {$result = "application/x-tex"}
                 "texi" {$result = "application/x-texinfo"}
                 "texinfo" {$result = "application/x-texinfo"}
                 "tgz" {$result = "application/x-compressed"}
                 "thmx" {$result = "application/vnd.ms-officetheme"}
                 "thn" {$result = "application/octet-stream"}
                 "tif" {$result = "image/tiff"}
                 "tiff" {$result = "image/tiff"}
                 "toc" {$result = "application/octet-stream"}
                 "tr" {$result = "application/x-troff"}
                 "trm" {$result = "application/x-msterminal"}
                 "tsv" {$result = "text/tab-separated-values"}
                 "ttf" {$result = "application/octet-stream"}
                 "txt" {$result = "text/plain"}
                 "u32" {$result = "application/octet-stream"}
                 "uls" {$result = "text/iuls"}
                 "ustar" {$result = "application/x-ustar"}
                 "vbs" {$result = "text/vbscript"}
                 "vcf" {$result = "text/x-vcard"}
                 "vcs" {$result = "text/plain"}
                 "vdx" {$result = "application/vnd.ms-visio.viewer"}
                 "vml" {$result = "text/xml"}
                 "vsd" {$result = "application/vnd.visio"}
                 "vss" {$result = "application/vnd.visio"}
                 "vst" {$result = "application/vnd.visio"}
                 "vsw" {$result = "application/vnd.visio"}
                 "vsx" {$result = "application/vnd.visio"}
                 "vtx" {$result = "application/vnd.visio"}
                 "vsto" {$result = "application/octet-stream"}
                 "wav" {$result = "audio/wav"}
                 "wax" {$result = "audio/x-ms-wax"}
                 "wbmp" {$result = "image/vnd.wap.wbmp"}
                 "wcm" {$result = "application/vnd.ms-works"}
                 "wdb" {$result = "application/vnd.ms-works"}
                 "wks" {$result = "application/vnd.ms-works"}
                 "wm" {$result = "video/x-ms-wm"}
                 "wma" {$result = "audio/x-ms-wma"}
                 "wmd" {$result = "application/x-ms-wmd"}
                 "wmf" {$result = "application/x-msmetafile"}
                 "wml" {$result = "text/vnd.wap.wml"}
                 "wmlc" {$result = "application/vnd.wap.wmlc"}
                 "wmls" {$result = "text/vnd.wap.wmlscript"}
                 "wmlsc" {$result = "application/vnd.wap.wmlscriptc"}
                 "wmp" {$result = "video/x-ms-wmp"}
                 "wmv" {$result = "video/x-ms-wmv"}
                 "wmx" {$result = "video/x-ms-wmx"}
                 "wmz" {$result = "application/x-ms-wmz"}
                 "wps" {$result = "application/vnd.ms-works"}
                 "wri" {$result = "application/x-mswrite"}
                 "wrl" {$result = "x-world/x-vrml"}
                 "wrz" {$result = "x-world/x-vrml"}
                 "wsdl" {$result = "text/xml"}
                 "wvx" {$result = "video/x-ms-wvx"}
                 "x" {$result = "application/directx"}
                 "xaf" {$result = "x-world/x-vrml"}
                 "xaml" {$result = "application/xaml+xml"}
                 "xbap" {$result = "application/x-ms-xbap"}
                 "xbm" {$result = "image/x-xbitmap"}
                 "xdr" {$result = "text/plain"}
                 "xla" {$result = "application/vnd.ms-excel"}
                 "xlam" {$result = "application/vnd.ms-excel.addin.macroEnabled.12"}
                 "xlc" {$result = "application/vnd.ms-excel"}
                 "xlm" {$result = "application/vnd.ms-excel"}
                 "xls" {$result = "application/vnd.ms-excel"}
                 "xlt" {$result = "application/vnd.ms-excel"}
                 "xlw" {$result = "application/vnd.ms-excel"}
                 "xlsb" {$result = "application/vnd.ms-excel.sheet.binary.macroEnabled.12"}
                 "xlsm" {$result = "application/vnd.ms-excel.sheet.macroEnabled.12"}
                 "xlsx" {$result = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
                 "xltm" {$result = "application/vnd.ms-excel.template.macroEnabled.12"}
                 "xltx" {$result = "application/vnd.openxmlformats-officedocument.spreadsheetml.template"}
                 "xml" {$result = "text/xml"}
                 "xof" {$result = "x-world/x-vrml"}
                 "xpm" {$result = "image/x-xpixmap"}
                 "xps" {$result = "application/vnd.ms-xpsdocument"}
                 "xsd" {$result = "text/xml"}
                 "xsf" {$result = "text/xml"}
                 "xsl" {$result = "text/xml"}
                 "xslt" {$result = "text/xml"}
                 "xsn" {$result = "application/octet-stream"}
                 "xtp" {$result = "application/octet-stream"}
                 "xwd" {$result = "image/x-xwindowdump"}
                 "z" {$result = "application/x-compress"}
                 "zip" {$result = "application/x-zip-compressed"}
                default {$result = "application/octet-stream"}
		}
	return $result
}

# Internal function used to translate PowerShell data types to SQL Server data types
$PStoSQLtypes = @{
    #PS datatype = SQL data type
    'System.Int32'    = 'int';
    'System.UInt32'   = 'bigint';
    'System.Int16'    = 'smallint';
    'System.UInt16'   = 'int';
    'System.Int64'    = 'bigint';
    'System.UInt64'   = 'decimal(20,0)';
    'System.Decimal'  = 'decimal(20,5)';
    'System.Single'   = 'bigint';
    'System.Double'   = 'float';
    'System.Byte'     = 'tinyint';
    'System.SByte'    = 'smallint';
    'System.TimeSpan' = 'nvarchar(30)';
    'System.String'   = 'nvarchar(MAX)';
    'System.Char'     = 'nvarchar(1)'
    'System.DateTime' = 'datetime2';
    'System.Boolean'  = 'bit';
    'System.Guid'     = 'uniqueidentifier';
    'Int32'           = 'int';
    'UInt32'          = 'bigint';
    'Int16'           = 'smallint';
    'UInt16'          = 'int';
    'Int64'           = 'bigint';
    'UInt64'          = 'decimal(20,0)';
    'Decimal'         = 'decimal(20,5)';
    'Single'          = 'bigint';
    'Double'          = 'float';
    'Byte'            = 'tinyint';
    'SByte'           = 'smallint';
    'TimeSpan'        = 'nvarchar(30)';
    'String'          = 'nvarchar(MAX)';
    'Char'            = 'nvarchar(1)'
    'DateTime'        = 'datetime2';
    'Boolean'         = 'bit';
    'Bool'            = 'bit';
    'Guid'            = 'uniqueidentifier';
    'int'             = 'int';
    'long'            = 'bigint';
}

##### EXPOSED FUNCTIONS #####

<#
 .Synopsis
  Retrieves token information for the Fluxx API.

 .Description
  Retrieves a JSON payload from the Fluxx API using the app ID and secret parameters. The
  bearer token in the payload is used for other Fluxx API calls

 .Parameter BaseURL
  The URL for the Fluxx instance being called.

 .Parameter ApplicationID
  The application ID assigned when registering an OAuth application at https://<your site>.fluxx.io/oauth/applications.

 .Parameter Secret
  The secret assigned when registering an OAuth application at https://<your site>.fluxx.io/oauth/applications.

 .Example
   # Retrieve the bearer token payload.
   Get-FluxxBearerToken -BaseUrl "<your site>.fluxx.io" -ApplicationID "<api id>" -Secret "<secret>"
#>
function Get-FluxxBearerToken {
param(
      [Parameter(Mandatory=$true)][String]$BaseUrl
    , [Parameter(Mandatory=$true)][String]$ApplicationID
    , [Parameter(Mandatory=$true)][String]$Secret
     )

    if(!$BaseUrl) { return Log-Message "Get-FluxxBearerToken" "ERROR: BaseURL is a required paramater" }
    if(!$ApplicationID) { return Log-Message "Get-FluxxBearerToken" "ERROR: ApplicationId is a required paramater" }
    if(!$Secret) { return Log-Message "Get-FluxxBearerToken" "ERROR: Secret is a required paramater" }

	$encodedClientId = [System.Web.HttpUtility]::UrlEncode($ApplicationID)
	$encodedClientSecret = [System.Web.HttpUtility]::UrlEncode($Secret)
	$uri = ([string]::Format("https://{0}/oauth/token?grant_type=client_credentials&client_id={1}&client_secret={2}", $BaseUrl, $encodedClientId, $encodedClientSecret))
    $response = $null
    try {
	    Log-Message "Get-FluxxBearerToken" "Retrieving Bearer Token"
         $response = Invoke-RestMethod -Uri $uri -Method POST
    } catch {
	    Log-Message "Get-FluxxBearerToken" "Token Retrieval Failed - Waiting 30 Seconds Before Attempting Again"
        Start-Sleep -Seconds 30
	    Log-Message "Get-FluxxBearerToken" "Retrieving Bearer Token"
        $response = Invoke-RestMethod -Uri $uri -Method POST
    }
	return $response
}
Export-ModuleMember -Function Get-FluxxBearerToken

<#
 .Synopsis
  Retrieves an object or list of objects via the Fluxx API.

 .Description
  Retrieves an object or list of objects via the Fluxx API.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  The name of the object to be returned via the API (e.g. grant_request).

 .Parameter QuerystringParameters
  [Optional] Used to override and add querystring parameters. It should begin with the parameters used to define which columns to return. If filters
  need to be applied, use this parameter. Defaults to "all_core=1&all_dynamic=1"

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Rertieve the first 25 grant_request records via the API. Returns an object including an array of records
   Get-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "grant_request"

 .Example
   # Rertieve a specific organization record via the API. Returns an object for a specific organization
   Get-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "organization/<org id>"

 .Example
   # Rertieve the core fields of the first 100 request_transaction records due within the next 3 months that haven't been paid. Returns an object including an array of records
   Get-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "request_transaction" -QuerystringParameters 'all_core=1&per_page=100&filter={"group_type":"and","conditions":[["due_at","next-n-months","3"],["paid_at","null","-"]]}'

#>
function Get-FluxxObject {
param(
      [Parameter(Mandatory=$true)][String]$BearerToken
    , [Parameter(Mandatory=$true)][String]$BaseUrl
    , [Parameter(Mandatory=$true)][String]$FluxxObject
    , [Parameter(Mandatory=$false)][String]$QuerystringParameters
    , [Parameter(Mandatory=$false)][String]$ApiVersion
     )

    if(!$ApiVersion) { $ApiVersion = "v2" }
    if(!$QuerystringParameters) {
        if($ApiVersion -eq "v1") { $QuerystringParameters = "style=full" }
        else { $QuerystringParameters = "all_core=1&all_dynamic=1" }
    } else {
        if($ApiVersion -eq "v1" -and !$QuerystringParameters.Contains("style=")) { $QuerystringParameters += "&style=full" }
        elseif(!$QuerystringParameters.Contains("all_") -and !$QuerystringParameters.Contains("cols=")) { $QuerystringParameters += "&all_core=1&all_dynamic=1" }
    }

	$fluxxUri = "https://{0}/api/rest/{1}/{2}?{3}" -f $BaseUrl, $ApiVersion, $FluxxObject, $QuerystringParameters
	Log-Message "Get-FluxxObject" $fluxxUri
	return Invoke-RestMethod -uri $fluxxUri -method GET -Headers @{Authorization=("Bearer {0}" -f $BearerToken)}
}
Export-ModuleMember -Function Get-FluxxObject


<#
 .Synopsis
  Creates a new object via the Fluxx API.

 .Description
  Creates a new object via the Fluxx API.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  The name of the object to be created via the API (e.g. grant_request).

 .Parameter Data
  The data to be used to create the new object made up of name and value pairs using the following syntax:
  @{
    name = 'A Sub Program'
    description = 'The Sub Program Description'
    program_id = 1
  }

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Create a new sub_program record via the API. Returns the newly created object
   $subprogram = @{
      name = 'A Sub Program'
      description = 'The Sub Program Description'
      program_id = 1
   }
   New-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "sub_program" -Data $subprogram
#>
function New-FluxxObject {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$true)][String]$FluxxObject
     , [Parameter(Mandatory=$true)][PSCustomObject]$Data
     , [Parameter(Mandatory=$false)][String]$ApiVersion
      )

    if(!$ApiVersion) { $ApiVersion = "v2" }
 
    $columns = ''
    foreach($key in $Data.keys) {
        if($columns -ne '' ) { $columns += ',' }
        else { $columns = '[' }
        $columns += ('"' + $key + '"')
    }
    if($Data.Keys -notcontains "id") { $columns += ',"id"' }
    $columns += ']'

 	$body = @{
              data = $Data | ConvertTo-Json
              cols = $columns
             }

	$fluxxUri = "https://{0}/api/rest/{1}/{2}" -f $BaseUrl, $ApiVersion, $FluxxObject
	Log-Message "New-FluxxObject" $fluxxUri
	return Invoke-RestMethod -uri $fluxxUri -Body $body -method POST -ContentType "multipart/form-data" -Headers @{Authorization=("Bearer {0}" -f $BearerToken)}
}
Export-ModuleMember -Function New-FluxxObject

<#
 .Synopsis
  Updates an existing object via the Fluxx API.

 .Description
  Updates an existing object via the Fluxx API.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  The name of the object to be update via the API (e.g. grant_request).

 .Parameter RecordID
  The ID of the record to be updated. The value should be an integer

 .Parameter Data
  The data to be used to create the new object made up of name and value pairs using the following syntax:
  @{
    name = 'An Existing Program'
    description = 'Updated Program Description'
  }

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Update an existing new sub_program record via the API. Returns the newly created object
   $subprogram = @{
      name = 'An Existing Sub Program'
      description = 'The Updated Sub Program Description'
   }
   Set-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "sub_program" -RecordID 12345 -Data $subprogram
#>
function Set-FluxxObject {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$true)][String]$FluxxObject
     , [Parameter(Mandatory=$true)][int]$RecordID
     , [Parameter(Mandatory=$true)][PSCustomObject]$Data
     , [Parameter(Mandatory=$false)][String]$ApiVersion)

    if(!$ApiVersion) { $ApiVersion = "v2" }
 
    $columns = ''
    foreach($key in $Data.keys) {
        if($columns -ne '' ) { $columns += ',' }
        else { $columns = '[' }
        $columns += ('"' + $key + '"')
    }
    if($Data.Keys -notcontains "id") { $columns += ',"id"' }
    $columns += ']'

 	$body = @{
              data = $Data | ConvertTo-Json
              cols = $columns
             }
	$fluxxUri = "https://{0}/api/rest/{1}/{2}/{3}" -f $baseUrl, $ApiVersion, $FluxxObject, $RecordID
	Log-Message "Set-FluxxObject" $fluxxUri
	return Invoke-RestMethod -uri $fluxxUri -Body $body -method PUT -ContentType "multipart/form-data" -Headers @{Authorization=("Bearer {0}" -f $BearerToken)}
}
Export-ModuleMember -Function Set-FluxxObject

<#
 .Synopsis
  Updates an existing object via the Fluxx API.

 .Description
  Updates an existing object via the Fluxx API.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  The name of the object to be update via the API (e.g. grant_request).

 .Parameter RecordID
  The ID of the record to be updated. The value should be an integer

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Delete an existing sub_program record via the API. Returns TRUE or FALSE depending on success
   Remove-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "sub_program" -RecordID 12345
#>
function Remove-FluxxObject {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$true)][String]$FluxxObject
     , [Parameter(Mandatory=$true)][int]$RecordID
     , [Parameter(Mandatory=$false)][String]$ApiVersion
      )
      
    if(!$ApiVersion) { $ApiVersion = "v2" }
	$fluxxUri = "https://{0}/api/rest/{1}/{2}/{3}" -f $BaseUrl, $apiVersion, $FluxxObject, $RecordID
	Log-Message "Remove-FluxxObject" $fluxxUri
    return Invoke-RestMethod -uri $fluxxUri -method DELETE -Headers @{Authorization=("Bearer {0}" -f $BearerToken)}
}
Export-ModuleMember -Function Remove-FluxxObject

<#
 .Synopsis
  Retrieves the entire list of a given Fluxx object.

 .Description
  This function uses the Fluxx API to iterate through all pages of a given Fluxx object type and returns an array of those objects.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  The name of the object to be update via the API (e.g. grant_request).

 .Parameter RecordsPerPage
  [Optional] The nuber of records to be returned per page. The default is 100 but can be increased up to 500. A larger number reduces the number of calls required
  to retrieve the full list but can have performance impacts for objects with a large number of attributes.

 .Parameter QuerystringParameters
  [Optional] Used to override and add querystring parameters. It should begin with the parameters used to define which columns to return. If filters
  need to be applied, use this parameter. Defaults to "all_core=1&all_dynamic=1"

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Retrieve all initiatives stored within Fluxx pulling 500 records at a time
   Get-FluxxObjectList -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "initiative" -RecordsPerPage 500

 .Example
   # Retrieve all grant_request records stored within Fluxx without a workflow state of declined
   Get-FluxxObjectList -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "grant_request" -QueryStringParameters 'filter={"group_type":"or","conditions":[["state","not-eq","declined"]]}'
#>
function Get-FluxxObjectList {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$true)][String]$FluxxObject
     , [Parameter(Mandatory=$false)][int]$RecordsPerPage
     , [Parameter(Mandatory=$false)][String]$QuerystringParameters
     , [Parameter(Mandatory=$false)][String]$ApiVersion
      )
      
    if(!$ApiVersion) { $ApiVersion = "v2" }
    if($RecordsPerPage -eq $null -or $RecordsPerPage -eq 0){ $RecordsPerPage = 100 }

    if($QuerystringParameters.Trim().Length -eq 0) { $QuerystringParameters = 'per_page=' + $RecordsPerPage }
    if(!$QuerystringParameters.Contains("per_page")) { $QuerystringParameters += '&per_page=' + $RecordsPerPage }

	Log-Message "Get-FluxxObjectList" ("Retrieving List from https://{0}/api/rest/{1}/{2}" -f $BaseUrl, $ApiVersion, $FluxxObject)
    $currentPage = 0
	$allRecords = @()
    $showTotals = $true
	do {
        $resp = Get-FluxxObject -BearerToken $BearerToken -ApiVersion $ApiVersion -FluxxObject $FluxxObject -BaseUrl $BaseUrl -QuerystringParameters ($QuerystringParameters + '&page=' + ($currentPage+1))
        if($showTotals) {
          Log-Message "Get-FluxxObjectList" ("Total Records: {0}" -f $resp.total_entries)
          Log-Message "Get-FluxxObjectList" ("Total Pages:   {0}" -f $resp.total_pages)
          $showTotals = $false
        }
        $ObjectName = ($resp.records | gm | where {$_.MemberType -eq 'NoteProperty'}).name
		$resp.records | foreach { $allRecords += $_.$ObjectName }
		$currentPage++
	} while ($currentPage -lt $resp.total_pages)
	return ,$allRecords
}
Export-ModuleMember -Function Get-FluxxObjectList

<#
 .Synopsis
  Downloads a specific document via the Fluxx API.

 .Description
  Downloads a specific document via the Fluxx API.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  [Optional] The name of the object to be update via the API. Defaults to "model_document"

 .Parameter DocumentID
  The ID of the document to be retrieved. The value should be an integer

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Retrieve a document with the ID of 1026325 via the Fluxx API
   Export-FluxxDocument -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -DocumentID 1026325
#>
function Export-FluxxDocument {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$false)][String]$FluxxObject
     , [Parameter(Mandatory=$true)][int]$DocumentID
     , [Parameter(Mandatory=$false)][String]$ApiVersion)
      
    if(!$ApiVersion) { $ApiVersion = "v2" }
    if(!$FluxxObject) { $FluxxObject = "model_document" }

	Log-Message "Get-FluxxDocument" ("Retrieving Document with Id: {0}" -f $DocumentID);
	$webclient = New-Object System.Net.WebClient
	$webclient.Headers.Add("Authorization", "Bearer $bearerToken")
    Log-Message "Get-FluxxDocument" ("https://{0}/api/rest/{1}/{2}/{3}" -f $BaseUrl, $ApiVersion, $FluxxObject, $DocumentID)
	return $webclient.DownloadData(("https://{0}/api/rest/{1}/{2}/{3}" -f $BaseUrl, $ApiVersion, $FluxxObject, $DocumentID))
}
Export-ModuleMember -Function Export-FluxxDocument

<#
 .Synopsis
  Uploads a document via the Fluxx API.

 .Description
  Uploads a document via the Fluxx API.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  [Optional] The name of the object to be update via the API. Defaults to "model_document"

 .Parameter FileName
  The full path of the file to be uploaded into Fluxx (e.g. C:\files\file-to-upload.txt)

 .Parameter ModelType
  The model type of the record associated with this file (e.g. grant_request)

 .Parameter ModelTypeOwnerID
  The ID of the model type of the record associated with this file

 .Parameter ModelTypeID
  The ID of the model_document_type associated with the file to be uploaded

 .Parameter UserID
  The ID of the people record associated with the upload

 .Parameter Description
  A description of the file being uploaded

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Example
   # Retrieve a document with the ID of 1026325 via the Fluxx API
   Import-FluxxDocument -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -DocumentID 1026325
#>
function Import-FluxxDocument {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$false)][String]$FluxxObject
     , [Parameter(Mandatory=$true)][String][String]$FileName
     , [Parameter(Mandatory=$true)][String][String]$ModelType
     , [Parameter(Mandatory=$true)][String][int]$ModelTypeOwnerID
     , [Parameter(Mandatory=$true)][String][int]$ModelTypeID
     , [Parameter(Mandatory=$true)][String][int]$UserID
     , [Parameter(Mandatory=$true)][String][String]$Description
     , [Parameter(Mandatory=$false)][String]$ApiVersion
      )
      
    if(!$ApiVersion) { $ApiVersion = "v2" }
    if(!$FluxxObject) { $FluxxObject = "model_document" }

	$contentType = Get-MimeContentType $FileName.substring($FileName.lastindexof('.')+1)
	$fileNameNoPath = $FileName.substring($FileName.lastindexof('\')+1)

	#Process Documents
	$documentParameterHash = @{}
	$documentParameterHash.Add("doc_label", "default")
	$documentParameterHash.Add("description", $Description)
	$documentParameterHash.Add("model_document_type", @{"id"=$ModelTypeID})
	$documentParameterHash.Add("owner_model", @{"model_type"="$ModelType";"id"=$ModelTypeOwnerID})
	$documentParameterHash.Add("storage_type", "file")
	$documentParameterHash.Add("content_type", $contentType)
	$documentParameterHash.Add("file_name", $fileNameNoPath)
	$documentParameterHash.Add("created_by_id", $UserID)
	$documentParameterHash.Add("updated_by_id", $UserID)
		
	$jsonString = $documentParameterHash | convertto-json -Compress

	#$jsonString
	
	$query = ("data={0}" -f [System.Web.HttpUtility]::UrlEncode($jsonString));
	$fileNameNoPath = $FileName.substring($FileName.lastindexof('\')+1)
	$boundary="-------------------------acebdf13572468"

    [System.Net.HttpWebRequest] $req = [System.Net.WebRequest]::create(("https://{0}/api/rest/{1}/{2}?{3}&style=full" -f $BaseUrl, $ApiVersion, $FluxxObject, $query))
    $req.Method = "POST"
    $req.ContentType = "multipart/form-data; boundary=$boundary"
	$req.Headers.Add("Authorization", "Bearer $BearerToken")
    $ContentLength = 0
    $req.TimeOut = 50000

    $reqst = $req.getRequestStream()

    $fileBuffer = [System.IO.File]::ReadAllBytes($FileName)

    <# part-header #>
    $header = "--$boundary`r`nContent-Disposition: form-data; name=`"content`"; filename=`"$fileNameNoPath`"`r`nContent-Type: $contentType`r`n`r`n"
    $buffer = [Text.Encoding]::ascii.getbytes($header)        
    $reqst.write($buffer, 0, $buffer.length)

    <# part-data #>
    $reqst.write($fileBuffer, 0, $fileBuffer.length)

    <# part-separator "One CRLF" #>
    $buffer = [Text.Encoding]::ascii.getbytes("`r`n")        
    $reqst.write($buffer, 0, $buffer.length)
 
    $buffer = [Text.Encoding]::ascii.getbytes("--$boundary--")        
    $reqst.write($buffer, 0, $buffer.length)
 
    $reqst.Flush()
    $reqst.Close()
    $reqst.Dispose()

    [net.httpWebResponse] $res = $req.getResponse()

    $resst = $res.getResponseStream()
    $sr = New-Object IO.StreamReader($resst)
    $result = $sr.ReadToEnd()
	
	$sr.Close()
	$sr.Dispose()

    $res.Close()
    $res.Dispose()
	$resst.Close()
    $resst.Dispose()

	return $result 
}
Export-ModuleMember -Function Import-FluxxDocument

<#
 .Synopsis
  Retrieves the entire list of a given Fluxx object and pushes it into SQL Server.

 .Description
  This function uses the Fluxx API to iterate through all pages of a given Fluxx object type and creates a table within SQL Server based on those records.

 .Parameter BearerToken
  A token used to authenticate against the Fluxx API. Can be retrieved using the access_token attribute retrieved from a Get-FluxxBearerToken call.

 .Parameter BaseUrl
  The URL for the Fluxx instance being called.

 .Parameter FluxxObject
  The name of the object to be update via the API (e.g. grant_request).

 .Parameter RecordsPerPage
  [Optional] The nuber of records to be returned per page. The default is 100 but can be increased up to 500. A larger number reduces the number of calls required
  to retrieve the full list but can have performance impacts for objects with a large number of attributes.

 .Parameter QuerystringParameters
  [Optional] Used to override and add querystring parameters. It should begin with the parameters used to define which columns to return. If filters
  need to be applied, use this parameter. Defaults to "all_core=1&all_dynamic=1"

 .Parameter ApiVersion
  [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

 .Parameter SQLServerName
  The full name of the SQL Server

 .Parameter SQLDatabaseName
  The name of the database in which the table will be created

 .Parameter SQLUserName
  The username to be used to connect to the SQL Server

 .Parameter SQLPassword
  The password to be used to connect to the SQL Server

 .Parameter SQLSchema
  [Optional] A schema name to be used. Defaults to fluxx

 .Parameter OverwriteTable
  [Optional] A switch used to to drop and recreate the table. If the switch isnt used, records will be appended to the table.

 .Example
   # Pull the grant_request records from the API and push them into SQL Server. If the table already exists, overwrite the existing table
   Export-FluxxObjectListToSQLServer -BearerToken  <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "grant_request" -SQLServerName <sql server name> -SQLDatabaseName <database name> -SQLUserName <database user name> -SQLPassword <database password> -OverwriteTable

#>
function Export-FluxxObjectListToSQLServer {
param (
       [Parameter(Mandatory=$true)][String]$BearerToken
     , [Parameter(Mandatory=$true)][String]$BaseUrl
     , [Parameter(Mandatory=$true)][String]$FluxxObject
     , [Parameter(Mandatory=$false)][int]$RecordsPerPage
     , [Parameter(Mandatory=$false)][String]$QuerystringParameters
     , [Parameter(Mandatory=$false)][String]$ApiVersion
     , [Parameter(Mandatory=$true)][String]$SQLServerName
     , [Parameter(Mandatory=$true)][String]$SQLDatabaseName
     , [Parameter(Mandatory=$true)][String]$SQLUserName
     , [Parameter(Mandatory=$true)][String]$SQLPassword
     , [Parameter(Mandatory=$false)][String]$SQLSchema
     , [Parameter(Mandatory=$false)][Switch]$OverwriteTable
      )

    $ProcessDate = Get-Date

    If(!$SQLSchema) { $SQLSchema = "fluxx" }
    try {
        Log-Message "Export-FluxxObjectListToSQLServer" "Retrieving $FluxxObject Records"
        $ObjectList = Get-FluxxObjectList $BearerToken $BaseUrl $FluxxObject $RecordsPerPage $QuerystringParameters $ApiVersion
    } catch {
        $Exception = $_.Exception
        Log-Message "Export-FluxxObjectListToSQLServer" "Error Retrieving Data - Exiting: $Message"
        throw $Exception
        return $false
    }

    try {
        [reflection.assembly]::GetAssembly([Fluxx.GrantMaker.DataTableUtility])
        Log-Message "Export-FluxxObjectListToSQLServer" "Fluxx.GrantMaker.DataTableUtility: Loaded"
    } catch {
        Log-Message "Export-FluxxObjectListToSQLServer" "Fluxx.GrantMaker.DataTableUtility: Loading"
        Add-Type -ReferencedAssemblies $Assem -TypeDefinition $Source -Language CSharp
        Log-Message "Export-FluxxObjectListToSQLServer" "Fluxx.GrantMaker.DataTableUtility: Loaded"
    }

    try {
        Log-Message "Export-FluxxObjectListToSQLServer" "Converting $FluxxObject Records to DataTable"
        $DataTable = New-Object Data.DataTable
        [Fluxx.GrantMaker.DataTableUtility]::ProcessRecords(($ObjectList | ConvertTo-Json),[ref] $DataTable)

        # Identify Date Columns
        $DateColumns = @()
        $StringColumns = $DataTable.Columns | Where {$_.DataType -like '*String' -and $_.ColumnName -notlike '*_xml'}
        foreach($StringColumn in $StringColumns.ColumnName) {
            $IsDate = $true
            $ColumnValues = ($DataTable | where {[string]::IsNullOrEmpty($_.$StringColumn) -eq $false} | Select id, $StringColumn -First 10)
            if($ColumnValues -is [Array]) {
                foreach($ColumnValue in $ColumnValues.$StringColumn) {
                    $DateCheck = $ColumnValue -as [DateTime]
                    if(!$DateCheck) {
                        $IsDate = $false
                        break
                    }
                }
            } else {
                $DateCheck = ($DataTable | where {[string]::IsNullOrEmpty($_.$ColumnValues)} | Select $ColumnValue -First 1).$ColumnValues
                if(!$DateCheck) { $IsDate = $false; }
            }
            if($IsDate) { $DateColumns += $StringColumn; }
        }
        
        Log-Message "Export-FluxxObjectListToSQLServer" "Formatting Date/Time Records within DataTable"
        [Fluxx.GrantMaker.DataTableUtility]::ProcessDates(($DateColumns -join ","),[ref] $DataTable)

        $ProcessDateColumn = New-Object system.Data.DataColumn api_import_date,([datetime])
        $ProcessDateColumn.DefaultValue = $ProcessDate
        $DataTable.Columns.Add($ProcessDateColumn)
    } catch {
        $Exception = $_.Exception
        Log-Message "Export-FluxxObjectListToSQLServer" "Error Generating DataTable - Exiting: $Message"
        throw $Exception
        return $false
    }

    Log-Message "Export-FluxxObjectListToSQLServer" "Establishing Connection to SQL Server"
    $TableName = $FluxxObject.ToLower()
    $ConnectionString = "Data Source=$SQLServerName;User ID=$SQLUserName;Password=$SQLPassword;Initial Catalog=$SQLDatabaseName"
    $Connection = new-object System.Data.SqlClient.SQLConnection($ConnectionString);
    try {
        $Connection.Open()
        Log-Message "Export-FluxxObjectListToSQLServer" "Connection Established"
    } catch {
        $Message = $_.Exception.Message
        Log-Message "Export-FluxxObjectListToSQLServer" "Error Establishing Connection - Exiting: $Message"
        return $false
    }
   
    Log-Message "Export-FluxxObjectListToSQLServer" "Checking if $SQLSchema.$TableName Exists"
    #CHECK IF TABLE EXISTS
    $SQLStatement = "SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[$SQLSchema].[$TableName]') AND type in (N'U')"
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand($SQLStatement,$Connection)
    try {
        $Result = $SQLCommand.ExecuteScalar()
    } catch {
        $Message = $_.Exception.Message
        Log-Message "Export-FluxxObjectListToSQLServer" "Table Check Falied - Exiting: $Message"
        $Connection.Close()
        return $false
    }
    If($Result) { Log-Message "Export-FluxxObjectListToSQLServer" "$SQLSchema.$TableName Exists" }

    #IF NOT, CREATE THE TABLE
    If(!$Result -or $OverwriteTable) {

        $sqldatatypes = @()
        foreach ($column in $DataTable.Columns) {
            $sqlcolumnname = $column.ColumnName
                    
            try {
                $columnvalue = $InputObject.Rows[0].$sqlcolumnname
            }
            catch {
                $columnvalue = $InputObject.$sqlcolumnname
            }
                    
            if ($columnvalue -eq $null) { $columnvalue = $InputObject.$sqlcolumnname }
                    
            
            if ($PStoSQLtypes.Keys -contains $column.DataType) {
                $sqldatatype = $PStoSQLtypes[$($column.DataType.toString())]
            }
            else {
                $sqldatatype = "nvarchar(MAX)"
            }
                    
            $sqldatatypes += "[$sqlcolumnname] $sqldatatype"
        }

        If($Result) { 
            Log-Message "Export-FluxxObjectListToSQLServer" "Dropping Table $SQLSchema.$TableName" 
            $SQLStatement = "BEGIN DROP TABLE $SQLSchema.$TableName END"
            $SQLCommand = New-Object System.Data.SqlClient.SqlCommand($SQLStatement,$Connection)
            try {
                $Result = $SQLCommand.ExecuteScalar()
            } catch {
                $Message = $_.Exception.Message
                Log-Message "Export-FluxxObjectListToSQLServer" "Dropping Table Failed - Exiting: $Message"
                $Connection.Close()
                return $false
            }
        }
        
        Log-Message "Export-FluxxObjectListToSQLServer" "Creating Table $SQLSchema.$TableName"
        $SQLStatement = "BEGIN CREATE TABLE $SQLSchema.$TableName ($($sqldatatypes -join ' NULL,')) END"
        $SQLCommand = New-Object System.Data.SqlClient.SqlCommand($SQLStatement,$Connection)
        try {
            $Result = $SQLCommand.ExecuteScalar()
        } catch {
            $Message = $_.Exception.Message
            Log-Message "Export-FluxxObjectListToSQLServer" "Table Creation Failed - Exiting: $Message"
            $Connection.Close()
            return $false
        }
    }
    
    Log-Message "Export-FluxxObjectListToSQLServer" "Loading Data Into $SQLSchema.$TableName"
    $BulkCopy = new-object ("System.Data.SqlClient.SqlBulkCopy") $Connection
    $BulkCopy.DestinationTableName = "$SQLSchema.$TableName"
    $BulkCopy.BulkCopyTimeout = 0
    try {
        $BulkCopy.WriteToServer($DataTable)
    } catch {
        $Message = $_.Exception.Message
        Log-Message "Export-FluxxObjectListToSQLServer" "DataTable Load Falied - Exiting: $Message"
        $Connection.Close()
        return $false
    }
    $Connection.Close()
    Log-Message "Export-FluxxObjectListToSQLServer" "Connection Closed"

	return $True
}
Export-ModuleMember -Function Export-FluxxObjectListToSQLServer

using System;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Net.Mail;
using System.Data.OleDb;
using System.Data.SqlClient;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using iTextSharp.text.xml.simpleparser;
using iTextSharp.text.html;
using System.IO;
using System.Collections.Specialized;

namespace BALCON_APP
{
    public class PDFGen
    {
        logcontrol lc = new logcontrol();

        // Connection string for the database
        string CONN_STRING_DB2 = "";
        // Variable to store error messages
        string error = string.Empty;
    
         // Main method to generate PDFs
        public string Pdf_gen(string cbs_username,string cbs_password)
        {
           
           // Read configuration settings from App.config

            string accountTypeValues = ConfigurationManager.AppSettings["Type"];
            string account_type1 = ConfigurationManager.AppSettings["Type"];            //read system config file
            string br_exclude = ConfigurationManager.AppSettings["Exclude"];

            string cbs = ConfigurationManager.AppSettings["CBS"];
            string cbs_unit = ConfigurationManager.AppSettings["Unit"];

           // string path = System.Configuration.ConfigurationManager.AppSettings["Path"];

            // Construct the connection string for the database

            CONN_STRING_DB2 = "Provider=IBMDA400.DataSource.1;Data Source=" + cbs + ";User ID=" + cbs_username + ";Password=" + cbs_password + ";Initial Catalog=" + cbs + ";Force Translate=37;Default Collection=*****";
            

            // Construct the unit and branch strings
            string unit = "KFIL" + cbs_unit;

            string br = br_exclude.Trim();
            string casa = accountTypeValues.Trim();

            try
            {
                // Query to get the range of records to process

                string Sql = "Select cntfrom,cntto from pdfrange";

                // Establish a connection to the database

                string constr1 = ConfigurationManager.ConnectionStrings["ConnectionString1"].ConnectionString;
                SqlConnection conn1 = new SqlConnection(constr1);

                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = Sql;
                cmd.Connection = conn1;
                conn1.Open();

                // Execute the query and read the range values
                SqlDataReader dr1 = null;
                dr1 = cmd.ExecuteReader();

                int From = 0;
                int To = 0;


                while (dr1.Read())
                {

                    From = Convert.ToInt32(dr1["cntfrom"]);
                    To = Convert.ToInt32(dr1["cntto"]);
                }
                // Close the connection
                conn1.Close();
                conn1.Dispose();

                // Construct the main SQL query to fetch account details
                string strSql = "";


                strSql = "Select CHAR(SUBSTR(A.EQDATE,6,2)|| '/' || SUBSTR(A.EQDATE, 4, 2) || '/' || '20' || SUBSTR(A.EQDATE, 2, 2)) AS EQ_DATE, A.CABRNM AS BRANCH, "
                               + "A.INTERNAL,CASE WHEN (A.EXTERNAL<>'') THEN A.EXTERNAL ELSE A.INTERNAL END EXTERNAL,A.GFCUN AS AC_NAME,A.ACCTYPE AS AC_TYPE,DECIMAL (SCBAL/100,15,2) As Balance, "
                               + "'As on December end 2024 balance of your ABBL AC :  '||A.EXTERNAL||' is BDT '||' '||DECIMAL (SCBAL/100,15,2) ||'.' AS SMS, "
                               + "A.MOBILE AS Mobile_NO, A.eMail AS EMAIL,CABRN,CATPH,C5ATD,C8CUR,ADD1,ADD2,ADD3,ADD4,ADD5,SLNO "
                               + "FROM LIBNAME.TABLENAME A WHERE SLNO BETWEEN " + From + ""
                               + " AND " + To + " ORDER BY SLNO";



                // Establish a connection to the database using OleDb

                OleDbConnection cn = new OleDbConnection(CONN_STRING_DB2);

                OleDbCommand insert = new OleDbCommand();
                insert.CommandText = strSql;
                insert.Connection = cn;
                cn.Open();
                OleDbDataReader dr = null;
                dr = insert.ExecuteReader();


                // Check if any rows were returned
                if (dr.HasRows == false)
                {

                   error = "*** No Row found to create Pdf! ";
                   return error;
                }

                int i = 0;

                // Process each row and generate a PDF
                while (dr.Read())
                {

                    string seq = dr["SLNO"].ToString().PadLeft(6, '0');

                    int SL = Convert.ToInt32(dr["SLNO"]);
                    // lblpdf.Text = String.Empty;
                    string password = "";

                    // Clean up the email address
                    string str = dr["EMAIL"].ToString();

                    string newstr = str.Replace("'", " ");

                    string new_str = newstr.Replace(",", " ");  //replace , with space

                    string temp = new_str.Replace(" ", "");             //remove trailing space

                    string Email = temp.Trim();

                    // Extract account details
                    string Branch = dr["BRANCH"].ToString();
                    string SMS = dr["SMS"].ToString();
                    string Internal = dr["INTERNAL"].ToString();
                    string External = dr["EXTERNAL"].ToString();  //Account

                    string address1 = string.Empty;

                    string address2 = string.Empty;

                    string address3 = string.Empty;

                    string address4 = string.Empty;

                    string address5 = string.Empty;

                    string branch = string.Empty;

                    string phone = string.Empty;

                    string account = string.Empty;

                    string accountno = string.Empty;

                    string currency = string.Empty;

                    branch = dr["CABRN"].ToString();
                    phone = dr["CATPH"].ToString();
                    account = dr["C5ATD"].ToString();
                    //     accountno = dr["ACCOUNT"].ToString();
                    currency = dr["C8CUR"].ToString();

                    address1 = dr["ADD1"].ToString();
                    address2 = dr["ADD2"].ToString();
                    address3 = dr["ADD3"].ToString();
                    address4 = dr["ADD4"].ToString();
                    address5 = dr["ADD5"].ToString();

                    // Create a new PDF document

                    Document doc = new Document(PageSize.A4);

                    string fileName = "BC" + Internal.Substring(0, 4) + "-" + DateTime.Now.ToString("yyMMdd") + "-" + seq + ".pdf";

                    string path = System.Configuration.ConfigurationManager.AppSettings["pathpdf"] + fileName;


                    //   var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(path.ToString(), FileMode.Create));
                    try
                    {
                        // Initialize the PDF writer
                        var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(path, FileMode.Create));
                    
                    doc.Open();

                    //-------------Image Water Mark-----------------------------

                    // Add a watermark image to the PDF

                    PdfContentByte cb = pdfWriter.DirectContentUnder;

                    //string relativePath = "\\Image\\WMlogo.png";
                    string relativePath = System.Configuration.ConfigurationManager.AppSettings["relativePath"]; ;
                    string imageWM_path = System.Configuration.ConfigurationManager.AppSettings["imageWM_path"];
                    string imageWM = Path.Combine(imageWM_path, relativePath);

                    iTextSharp.text.Image imgWM = iTextSharp.text.Image.GetInstance(imageWM);

                    imgWM.SetAbsolutePosition(50f, 350f);

                    imgWM.ScaleToFit(500, 1500);

                    PdfGState state = new PdfGState();
                    state.FillOpacity = 0.1f;
                    cb.SetGState(state);
                    cb.AddImage(imgWM);

                    //-------------Image Water Mark-----------------------------

                    // Add a logo to the PDF

                    string relativeURL = System.Configuration.ConfigurationManager.AppSettings["relativeURL"];
                    string imageURL_path = System.Configuration.ConfigurationManager.AppSettings["imageURL_path"];
                    string imageURL = Path.Combine(imageURL_path, relativeURL);

                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imageURL);
                    img.Alignment = Element.ALIGN_LEFT;

                    img.SetAbsolutePosition(55f, 760f);
                    img.SpacingAfter = 30f;
                    //img.ScaleAbsolute(250f, 65f);

                    img.ScalePercent(70f);
                    //img.ScaleToFit(150f, 60f);

                    doc.Add(img);

                    //===================================================================================
                     // Define fonts for the PDF content

                    Font Tahoma0 = FontFactory.GetFont("Tahoma", 8f, Font.BOLD, Color.BLACK);

                    Font Tahoma1 = FontFactory.GetFont("Tahoma", 9f, Font.NORMAL, Color.BLACK);

                    Font Tahoma2 = FontFactory.GetFont("Tahoma", 10f, Font.BOLD, new Color(204, 0, 0));

                    Font Tahoma3 = FontFactory.GetFont("Tahoma", 10f, Font.BOLD, Color.BLACK);

                    Font Tahoma4 = FontFactory.GetFont("Tahoma", 9f, Font.BOLD, new Color(0, 153, 0));

                    Font Tahoma5 = FontFactory.GetFont("Tahoma", 6f, Font.NORMAL, Color.BLACK);

                    Font Tahoma6 = FontFactory.GetFont("Tahoma", 6f, Font.BOLD, Color.BLACK);

                    Font Tahoma7 = FontFactory.GetFont("Tahoma", 6f, Font.NORMAL, Color.BLUE);

                    //--------------------------Table header----------------------------------
                    // Create a table for the address details

                    PdfPTable table = new PdfPTable(1);

                    table.HorizontalAlignment = 1;
                    table.TotalWidth = 300f;
                    table.LockedWidth = true;

                    PdfPCell Col11 = new PdfPCell(new Phrase(new Chunk(address1.ToString(), Tahoma0))); //Col 1 Row 1
                    Col11.Border = 0;
                    table.AddCell(Col11);

                    PdfPCell Col12 = new PdfPCell(new Phrase(new Chunk(address2.ToString(), Tahoma0))); //Col 1 row 2
                    Col12.Border = 0;
                    table.AddCell(Col12);

                    PdfPCell Col13 = new PdfPCell(new Phrase(new Chunk(address3.ToString(), Tahoma0)));
                    Col13.Border = 0;
       
                    table.AddCell(Col13);

                    PdfPCell Col14 = new PdfPCell(new Phrase(new Chunk(address4.ToString(), Tahoma0)));
                    Col14.Border = 0;
                    table.AddCell(Col14);

                    PdfPCell Col15 = new PdfPCell(new Phrase(new Chunk(address5.ToString(), Tahoma0))); //col5 row 1
                    Col15.Border = 0;
                    table.AddCell(Col15);

                    table.WriteSelectedRows(0, -1, 55, 750, pdfWriter.DirectContent);

                    Paragraph lineSeparator1 = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
                    // Set gap between line paragraphs.
                    lineSeparator1.SetLeading(0.5F, 0.5F);
                    lineSeparator1.IndentationLeft = 20f;

                    lineSeparator1.SpacingBefore = 115F;

                    Chunk c1 = new Chunk("Dear Valued Customer:", Tahoma1);

                    //Chunk c2 = new Chunk("Thank you for Banking with AB Bank Ltd.", Tahoma1);
                    Chunk c2 = new Chunk("We are grateful that you have chosen AB Bank for all your Banking needs.", Tahoma1);

                    Chunk c3 = new Chunk("Balance Confirmation", Tahoma2);
                    Chunk c4 = new Chunk("AB Bank PLC. firmly believes in preserving the environment and has taken a number of Green Banking "
                                       + " initiatives. You will be happy to know that in the spirit of green banking guidelines of Bangladesh bank,"
                                       + " we have commenced sending Balance Confirmation to our valued Customers through e-mail.", Tahoma1);

                    Chunk c5 = new Chunk(SMS.ToString(), Tahoma3);

                    Chunk c6 = new Chunk("Discontinue of paper Statement", Tahoma2);
                    Chunk c7 = new Chunk("You will no longer receive your Balance Confirmation of your Account(s) through paper statements.", Tahoma1);

                    Chunk c10 = new Chunk("LET US ALL JOIN HANDS AND SAVE OUR ENVIRONMENT.", Tahoma4);

                    Chunk F1 = new Chunk("This email and any files transmitted with it are confidential and intended solely for the use of the individual(s)"
                        + " or organization to whom they are addressed. If you are not the intended recipient(s) and have received this email in error, please notify ", Tahoma5);

                    Chunk F2 = new Chunk("support@abbl.com", Tahoma7);
                    F2.SetUnderline(0.3f, -2f);

                    Chunk F3 = new Chunk(" or ", Tahoma5);

                    Chunk F4 = new Chunk("call 16207", Tahoma6);

                    Chunk F5 = new Chunk(" or ", Tahoma5);

                    Chunk F6 = new Chunk("+8800000000000 ", Tahoma6);

                    Chunk F7 = new Chunk("immediately and please do not disseminate, distribute, print or copy this email or its contents or its attachment by any means.", Tahoma5);

                    Phrase ph1 = new Phrase();
                    Phrase ph2 = new Phrase();
                    Phrase ph3 = new Phrase();
                    Phrase ph4 = new Phrase();
                    Phrase ph5 = new Phrase();
                    Phrase ph6 = new Phrase();
                    Phrase ph7 = new Phrase();
                    Phrase ph10 = new Phrase();
                    Phrase Phf1 = new Phrase();

                    ph1.Add(c1);
                    ph2.Add(c2);
                    ph3.Add(c3);
                    ph4.Add(c4);
                    ph5.Add(c5);
                    ph6.Add(c6);
                    ph7.Add(c7);
                    ph10.Add(c10);
                    Phf1.Add(F1);
                    Phf1.Add(F2);
                    Phf1.Add(F3);
                    Phf1.Add(F4);
                    Phf1.Add(F5);
                    Phf1.Add(F6);
                    Phf1.Add(F7);

                    Paragraph p1 = new Paragraph();
                    p1.SpacingBefore = 30f;
                    p1.IndentationLeft = 20f;
                    p1.Add(ph1);

                    Paragraph p2 = new Paragraph();
                    p2.SpacingBefore = 20f;
                    p2.IndentationLeft = 20f;
                    p2.Add(ph2);

                    Paragraph p3 = new Paragraph();
                    p3.SpacingBefore = 20f;
                    p3.IndentationLeft = 20f;
                    p3.Add(ph3);

                    Paragraph p4 = new Paragraph();
                    p4.IndentationLeft = 20f;
                    p4.Add(ph4);

                    Paragraph p5 = new Paragraph();
                    p5.SpacingBefore = 20f;
                    p5.IndentationLeft = 20f;
                    p5.Add(ph5);

                    Paragraph p6 = new Paragraph();
                    p6.SpacingBefore = 20f;
                    p6.IndentationLeft = 20f;
                    p6.Add(ph6);

                    Paragraph p7 = new Paragraph();
                    p7.IndentationLeft = 20f;
                    p7.Add(ph7);

       
                    Paragraph p10 = new Paragraph();
                    p10.SpacingBefore = 20f;
                    p10.IndentationLeft = 20f;
                    p10.Add(ph10);

                    Paragraph lineSeparator = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 58.0F, new Color(0, 153, 0), Element.ALIGN_LEFT, 1)));
       
                    lineSeparator.SetLeading(0.5F, 0.5F);
                    lineSeparator.IndentationLeft = 20f;
                    lineSeparator.SpacingAfter = 70F;

                    Paragraph p13 = new Paragraph();
                    p13.SetLeading(0, 1.0f);
                    p13.IndentationLeft = 20f;

                    p13.Add(Phf1);

                    doc.Add(lineSeparator1);
                    doc.Add(p1);
                    doc.Add(p2);

                    doc.Add(p3);
                    doc.Add(p4);
                    doc.Add(p5);
                    doc.Add(p6);
                    doc.Add(p7);
                    doc.Add(p10);

                    // Footer
                    // Add footer to the PDF

                    Paragraph footerLineSep = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 2)));
                    // Set gap between line paragraphs.
                    footerLineSep.SetLeading(0.10F, 0.5F);

                    //  Paragraph footer = new Paragraph("Computer Generated Report", FontFactory.GetFont(FontFactory.TIMES_BOLD, 15, iTextSharp.text.Font.BOLD));
                    footerLineSep.Alignment = Element.ALIGN_LEFT;

                    PdfPTable footerTbl = new PdfPTable(1);
                    footerTbl.TotalWidth = 530;

                    footerTbl.DefaultCell.Border = Rectangle.NO_BORDER;

                    footerTbl.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
                    footerTbl.AddCell(footerLineSep);
                    footerTbl.AddCell(p13);
                    footerTbl.WriteSelectedRows(0, -1, 33, 60, pdfWriter.DirectContent); //15

                   // Close the document

                    doc.Close();

                    // Password-protect the PDF

                    //Comment    
                    string MFileName = string.Empty;
                    object TargetFile = path.ToString();

                    iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(path.ToString());

                    MFileName = System.Configuration.ConfigurationManager.AppSettings["pathpdf"] + "AB" + fileName.ToString();
                    // MFileName = MFileName.Insert


                    password = External.Substring(0, 6);

                    //password = Account.Substring(0, 4) + Account.Substring(5, 2);

                    iTextSharp.text.pdf.PdfEncryptor.Encrypt(reader, new FileStream(MFileName, FileMode.Append), iTextSharp.text.pdf.PdfWriter.STRENGTH128BITS, password.ToString().Trim(), "", iTextSharp.text.pdf.PdfWriter.AllowPrinting);
                    //Comment

                     // Log the generated PDF details into the database
                    string constr = ConfigurationManager.ConnectionStrings["ConnectionString1"].ConnectionString;
                    SqlConnection conn = new SqlConnection(constr);

                    conn.Open();
                    SqlCommand com = new SqlCommand();
                    com.Connection = conn;
                    com.CommandType = CommandType.Text;

                    DateTime dt = DateTime.Now;

                    com.CommandText = "insert into Attachment(Branch,Internal_AC,External_AC,Pdffile,Email,sms,Process_date,issent,seq,SL) values('" + Branch.ToString() + "','" + Internal.ToString() + "','" + External.ToString() + "','" + MFileName.ToString() + "','" + Email.ToString().Trim() + "','" + SMS.ToString() + "', '" + dt + "', 0,'" + seq.ToString() + "', " + SL + " )";
                    com.ExecuteNonQuery();

                    conn.Close();


                    ++i;

                    string seqcount = seq.ToString();

                    }
                    catch (Exception ex)
                    {
                         // Log any errors that occur during PDF generation
                        lc.Write("ERROR: " + ex.ToString());
                    }
                    //Comment         
                }



                // CTST=CTST.Substring(0)

                dr = null;
                //dr.Close();
                cn.Close();
                cn.Dispose();

                error = "*** All PDF Prepared To send e-mail! ";

            }

            catch (Exception ex)
            {
                // Handle any exceptions that occur during the process
                error = ex.ToString().Substring(0, 150);
                lc.Write("Final error:  "+ ex.ToString());
            }
            finally
            {
                 // Clear sensitive data
                cbs_username = String.Empty;
                cbs_password = String.Empty;
            }
            return error;
        }
    }
}

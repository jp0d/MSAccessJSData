//console.log("Working fine!");
//alert(1);

          
          function test()
            {
              console.log('Function begins!!');
            }
           
          
          function query()
          {
              var pad = "C:\\BackUp\\PowerBI\\js\\db\\test.accdb";
              var cn = new ActiveXObject("ADODB.Connection");
              var strConn = "Provider=microsoft.ace.oledb.12.0;Data Source=" + pad;
              cn.Open(strConn);
              var rs = new ActiveXObject("ADODB.Recordset");
              var SQL = "SELECT * FROM Tab1";
              rs.Open(SQL, cn);
              if(!rs.bof) 
              {
                  rs.MoveFirst();
                  while(!rs.eof)
                  {
                      document.write("<p>" + rs.fields(0).value + ", ");
                      document.write(rs.fields(1).value + ", ");
                      document.write(rs.fields(2).value + ".</p>");
                      rs.MoveNext();
                  }
              }
              else
              {
                  document.write("No data found");
              }
              rs.Close();
              cn.Close();
          }

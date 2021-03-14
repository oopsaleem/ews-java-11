package hi;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class Program {

    public static void main(String[] args) throws SQLException, ClassNotFoundException {
        Class.forName("com.cnsconnect.mgw.jdbc.MgDriver");
        // STEP 1: Create a JDBC connection to each target server
        // you get the data for the connection string from Connect Bridge
        String exchangeConnectionString = "jdbc:MgDriver:IMPL=CORBA;ENC=UTF8;HOST=123.456.789.000;PORT=8087;UID=demouser;PWD='password';ACC=accountExchange;";
        String sharepointConnectionString =        "jdbc:MgDriver:IMPL=CORBA;ENC=UTF8;HOST=123.456.789.000;PORT=8087;UID=demouser;PWD='password';ACC=accountSharePoint;";
        Connection exchangeConn =
                DriverManager.getConnection(exchangeConnectionString);
        Connection sharepointConn =
                DriverManager.getConnection(sharepointConnectionString);
        Statement exchangeSt = exchangeConn.createStatement();
        System.out.println("Connecting to Exchange...");
        //STEP 2: Provide an appropriate object like a ResultSet in JAVA
        //STEP 3: Fill the object with data from the source server
        ResultSet exchangeRs = exchangeSt.executeQuery("SELECT * FROM [Appointment]");
        //create a new JDBC statement for inserting PreparedStatement
        var sharepointSt = sharepointConn.prepareStatement("INSERT INTO [Calendar] ([Title], [Description], [Location], [StartTime], [EndTime]) VALUES ( ?, ?, ?, ?, ?)");
        //STEP 4: Manipulate the data and or apply a workflow rule
        //in this case, check if the appointment is private, if not, insert it into sharepoint
        while (exchangeRs.next()) {
            Boolean isPrivate = exchangeRs.getBoolean("IsPrivate");
            if (isPrivate != null && isPrivate)
            {
                System.out.println("Skipping '" + exchangeRs.getString("Subject") + "'");
                continue;
            }
            // Title
            //fill its parameters with values for the sharepoint account
            sharepointSt.setString(1, exchangeRs.getString("Subject"));
            sharepointSt.setString(2, exchangeRs.getString("Body"));
            sharepointSt.setString(3, exchangeRs.getString("Location"));
            sharepointSt.setTimestamp(4, exchangeRs.getTimestamp("StartDate"));
            sharepointSt.setTimestamp(5, exchangeRs.getTimestamp("EndDate"));

            System.out.println("Inserting '" + exchangeRs.getString("Subject") + "'");
            //STEP 5: INSERT the data into the target server
            sharepointSt.execute();
        }
        exchangeRs.close();
        exchangeSt.close();
        sharepointSt.close();
        //STEP 6: Close Connections
        exchangeConn.close();
        sharepointConn.close();
    }
}

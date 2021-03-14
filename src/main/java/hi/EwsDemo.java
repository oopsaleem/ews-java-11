package hi;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

import java.net.URI;

public class EwsDemo {
    public static void main(String[] args) {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials("sampleEmail", "securePa$$w0rd", "production");
        service.setCredentials(credentials);

        try {
            service.setUrl(new URI("https://webmail.email.com/owa/?bO=1#path=/mail"));
            service.autodiscoverUrl("sampleEmail@email.com");
        } catch (Exception e) {
            e.printStackTrace();
        }

        /*
        Folder folder = null;
        try {
            folder = new Folder(service);

            folder.setDisplayName("EWS-JAVA-Folder");
            folder.save(WellKnownFolderName.Inbox);
        } catch (Exception e) {
            e.printStackTrace();
        }
        */

        EmailMessage msg= null;
        try {
            msg = new EmailMessage(service);
            msg.setSubject("EWS Java API: Testing from Khaled -");
            MessageBody mb = MessageBody.getMessageBodyFromText("Sent using the <b>EWS Java</b> API. Amazing! <bre>");
            mb.setBodyType(BodyType.HTML);
            msg.setBody(mb);
            msg.getToRecipients().add("sampleEmail@email.com");
            msg.save(WellKnownFolderName.Drafts);
//            msg.send();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

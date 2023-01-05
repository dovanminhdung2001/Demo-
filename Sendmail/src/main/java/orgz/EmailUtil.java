package orgz;

import javax.mail.Message;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.util.Date;
import java.util.Properties;

public class EmailUtil {
    public static void sendEmail(String to, String subject, String msg, String user, String pass){
        Properties properties = new Properties();

        properties.put("mail.smtp.host", "smtp.gmail.com");
        properties.put("mail.smtp.port", "587");
        properties.put("mail.smtp.auth", true);
        properties.put("mail.smtp.starttls.enable", "true");

        Session session = Session.getDefaultInstance(properties, new javax.mail.Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(user, pass);
            }
        });

        try {
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(user));
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
            message.setSubject(subject);
            message.setContent(msg, "text/html");

            Transport.send(message);
            System.out.println("Done");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        String user = "14052001as@gmail.com";
        String pass = "rgegghzilrerzrgg";
        String to = "dovanminhdung2001@gmail.com";
        String subject = "test";
        String msg = "done";

        EmailUtil.sendEmail(to, subject, msg, user, pass);
    }
}

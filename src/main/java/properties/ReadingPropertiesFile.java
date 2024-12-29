package properties;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Collection;
import java.util.Properties;
import java.util.Set;

public class ReadingPropertiesFile {

    public static void main(String[] args) throws IOException {
        Properties properties = new Properties();
        FileInputStream fi = new FileInputStream(System.getProperty("user.dir")+"\\test_data\\config.properties");
        properties.load(fi);
        String url = properties.getProperty("appurl");
        String email = properties.getProperty("email");
        String password = properties.getProperty("password");
        String orid = properties.getProperty("orderid");
        String custid = properties.getProperty("customerid");

        System.out.println(url+"  "+email+"  "+password+"  "+orid+"  "+custid);

        //reading all the keys :
        Set<String> keys = properties.stringPropertyNames();
        System.out.println(keys);

        //reading all the values
        Collection<Object> value = properties.values();
        System.out.println(value);

        fi.close();
    }
}

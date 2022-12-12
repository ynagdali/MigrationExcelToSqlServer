package config;


import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.log4j.Logger;




public class ApplicationProperties {

    private final Properties properties;
    private static Logger logger = Logger.getLogger(ApplicationProperties.class);

    public ApplicationProperties() {
        properties = new Properties();
        InputStream input = null;
        try {
        	input = new FileInputStream("application.properties");
            properties.load(input);
            
        } catch (IOException ioex) {
        	ioex.printStackTrace();
        	logger.debug("IOException Occured while loading properties file::::" +ioex.getMessage());
        }
    }

    public String getProperty(String keyName) {
    	logger.debug("Reading Property " + keyName);
        return properties.getProperty(keyName, "There is no key in the properties file");
    }

}
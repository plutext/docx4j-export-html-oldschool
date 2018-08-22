package org.docx4j.convert.out.html;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.docx4j.utils.ResourceUtils;


public class ConverterProperties {
	
	protected static Logger log = Logger.getLogger(ConverterProperties.class);
	
	private static Properties properties;
	
	private static void init() {
		
		properties = new Properties();
		try {
			properties.load(
					ResourceUtils.getResource("converter.properties"));
		} catch (Exception e) {
			log.error("Error reading converter.properties", e);
		}
	}
	
	public static String getProperty(String name) {
		
		if (properties==null) {init();}
		return properties.getProperty(name);		
	}

	public static Properties getProperties() {
		
		if (properties==null) {init();}
		return properties;		
	}
	
}

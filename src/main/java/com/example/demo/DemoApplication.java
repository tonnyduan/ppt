package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.ConfigurationProperties;

@SpringBootApplication
@ConfigurationProperties
public class DemoApplication {

	public static void main(String[] args) throws  Exception{
		SpringApplication.run(DemoApplication.class, args);
	}

}

package com.arcanepost.service.spreadsheetutil;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@ComponentScan(basePackageClasses = SpreadsheetUtilController.class)
public class SpreadsheetUtilApplication extends SpringBootServletInitializer {

	@Override
	protected SpringApplicationBuilder configure(SpringApplicationBuilder application){
		return application.sources(SpreadsheetUtilController.class);
	}

	public static void main(String[] args) {

		SpringApplication.run(SpreadsheetUtilApplication.class, args);
	}
}

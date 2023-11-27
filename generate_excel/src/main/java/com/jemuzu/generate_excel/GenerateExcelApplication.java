package com.jemuzu.generate_excel;

import java.nio.file.Paths;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.Banner;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.jemuzu.generate_excel.process.GenerateExcelProcess;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@SpringBootApplication
public class GenerateExcelApplication implements CommandLineRunner {

	@Autowired
	GenerateExcelProcess generateExcelProcess;

	public static void main(String[] args) {
		SpringApplication app = new SpringApplication(GenerateExcelApplication.class);
		app.setBannerMode(Banner.Mode.OFF);
		app.run(args);

	}

	@Override
	public void run(String... args) throws Exception {

		if (args.length == 0) {
			String sourcePath = Paths.get(GenerateExcelApplication.class.getResource("/sample data.json").toURI()).toFile().getAbsolutePath();
			String targetPath = "./";
			String[] a = { "-s", sourcePath, "-t", targetPath };
			args = a;
		}

		CommandLine cmd = parseArguments(args);

		log.debug("Source Path: " + cmd.getOptionValue("s")); // file with data to write in excel
		log.debug("Target Path: " + cmd.getOptionValue("t")); // directory to save the excel file

		generateExcelProcess.executeProcess(cmd.getOptionValue("s"), cmd.getOptionValue("t"));

	}

	/**
	 * Parse the arguments
	 * 
	 * @param args
	 * @return CommandLine
	 */
	private static CommandLine parseArguments(String[] args) {
		Options options = new Options();

		Option source = new Option("s", "source", true, "Source Path - .csv or .json(json array)");
		source.setRequired(true);
		options.addOption(source);

		Option target = new Option("t", "target", true, "Target Path");
		target.setRequired(true);
		options.addOption(target);

		HelpFormatter formatter = new HelpFormatter();
		CommandLine cmd = null;

		try {
			cmd = new DefaultParser().parse(options, args);
		} catch (Exception e) {
			log.error(e.getMessage());
			formatter.printHelp("Pass parameters", options);
			System.exit(1);
		}
		return cmd;
	}

}
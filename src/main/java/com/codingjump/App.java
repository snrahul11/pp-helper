package com.codingjump;

import java.io.FileOutputStream;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class App {
    public static CommandLine parseArguments(String[] args) throws ParseException {
        HelpFormatter formatter = new HelpFormatter();

        Options options = new Options();
        options.addOption("e", "export", true, "Exports powerpoint file with given name");

        formatter.printHelp("pp-helper", options);

        CommandLineParser parser = new DefaultParser();
        CommandLine cmd;

        cmd = parser.parse(options, args);
        return cmd;
    }

    public static void main(String[] args) throws Exception {
        var cmd = parseArguments(args);

        try (var ppt = new XMLSlideShow()) {
            ppt.createSlide();

            FileOutputStream out;
            if (cmd.hasOption("export"))
                out = new FileOutputStream(cmd.getOptionValue("export") + ".pptx");
            else
                out = new FileOutputStream("output.pptx");
            ppt.write(out);
            out.close();
        }
    }
}

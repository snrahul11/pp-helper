package com.codingjump;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Collection;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.DirectoryFileFilter;
import org.apache.commons.io.filefilter.RegexFileFilter;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;

public class App {
    static Options options;

    static String exportOption = "export";

    public static CommandLine parseArguments(String[] args) throws ParseException {
        options = new Options();
        options.addOption("e", exportOption, true, "Exports powerpoint file with given name");
        options.addOption("d", "dir", true, "Location of directory where images are stored");
        options.addOption("r", "run", false, "If argment present then it will execute otherwise not");

        CommandLineParser parser = new DefaultParser();
        CommandLine cmd;

        cmd = parser.parse(options, args);
        return cmd;
    }

    public static void main(String[] args) throws Exception {
        var cmd = parseArguments(args);

        String outFileLocation;
        String dirLocation;
        boolean run = false;
        if (cmd.hasOption(exportOption))
            outFileLocation = cmd.getOptionValue("export") + ".pptx";
        else
            outFileLocation = "output.pptx";

        if (cmd.hasOption("dir"))
            dirLocation = cmd.getOptionValue("dir");
        else
            dirLocation = "./";

        if (cmd.hasOption("run"))
            run = true;

        if (!run) {
            HelpFormatter formatter = new HelpFormatter();
            formatter.printHelp("pp-helper", options);
            return;
        }

        try (var ppt = new XMLSlideShow()) {

            File dir;
            try (FileOutputStream outFile = new FileOutputStream(outFileLocation)) {
                dir = new File(dirLocation);

                Collection<File> files = FileUtils.listFiles(
                        dir,
                        new RegexFileFilter("^(.*?)(.jpg|.png|.svg|.jpeg|.gif)"),
                        DirectoryFileFilter.DIRECTORY);

                // Files imported
                for (var file : files) {
                    var slide = ppt.createSlide();
                    byte[] pictureData = IOUtils.toByteArray(new FileInputStream(file));
                    XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                    var picture = slide.createPicture(pd);
                    var dimensions = slide.getSlideShow().getPageSize();
                    var anchor = new Rectangle2D.Float(0f, 0f, (float) dimensions.getWidth(),
                            (float) dimensions.getHeight());
                    picture.setAnchor(anchor);
                    System.out.println("Image imported: " + file);
                }

                ppt.write(outFile);
                // Exported file
                System.out.println("Exported file: " + outFileLocation);
            }
        }
    }
}

package com.codingjump;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.filefilter.DirectoryFileFilter;
import org.apache.commons.io.filefilter.RegexFileFilter;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class App implements CommandLineRunner {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

    static Options options;

    static String ppOption = "powerpoint";
    static String importImagesOption = "import-images";
    static String exportImagesOption = "export-images";
    static String dirOption = "dir";

    public static CommandLine parseArguments(String[] args) throws ParseException {
        options = new Options();
        options.addOption("pp", ppOption, true, "Process powerpoint file with given name");
        options.addOption("d", dirOption, true, "Location of directory where images are stored");
        options.addOption("ii", importImagesOption, false, "Will import images to the PPT");
        options.addOption("ei", exportImagesOption, false, "Will export images from the PPT");

        CommandLineParser parser = new DefaultParser();
        CommandLine cmd;

        cmd = parser.parse(options, args);
        return cmd;
    }

    public static void main(String[] args) {
        SpringApplication.run(App.class, args);
    }

    private static void resizeToDimensions(XSLFPictureShape picture, Dimension dimensions) {
        var originalAnchor = picture.getAnchor();

        double h = originalAnchor.getHeight();
        double w = originalAnchor.getWidth();

        double mh = dimensions.getHeight();
        double mw = dimensions.getWidth();

        double rh = h;
        double rw = w;

        if (h > mh) {
            double p = rh;
            rh = mh;
            rw = rw * mh / p;
        }

        if (rw > mw) {
            double p = rw;
            rw = mw;
            rh = rh * mw / p;
        }

        originalAnchor.setFrame(0, 0, rw, rh);
        picture.setAnchor(originalAnchor);
    }

    @Override
    public void run(String... args) throws Exception {
        var cmd = parseArguments(args);

        String ppLocation;
        File dirLocation;

        if (cmd.hasOption(dirOption))
            dirLocation = new File(cmd.getOptionValue(dirOption));
        else
            dirLocation = new File("./");

        if (cmd.hasOption(ppOption))
            ppLocation = cmd.getOptionValue(ppOption);
        else
            ppLocation = "output.pptx";

        if (cmd.hasOption(importImagesOption)) {
            importImagesToNewPPT(ppLocation, dirLocation);
            return;
        } else if (cmd.hasOption(exportImagesOption)) {
            exportImagesFromPPT(ppLocation, dirLocation);
            return;
        }

        HelpFormatter formatter = new HelpFormatter();
        formatter.printHelp("pp-helper", options);
    }

    private void exportImagesFromPPT(String inFileLocation, File dirLocation)
            throws IOException {
        try (FileInputStream inFile = new FileInputStream(inFileLocation)) {
            try (var ppt = new XMLSlideShow(inFile)) {
                dirLocation.mkdirs();
                for (var slide : ppt.getSlides()) {
                    var dimensions = slide.getSlideShow().getPageSize();
                    BufferedImage img = new BufferedImage(dimensions.width, dimensions.height,
                            BufferedImage.TYPE_INT_RGB);
                    Graphics2D graphics = img.createGraphics();
                    graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
                            RenderingHints.VALUE_INTERPOLATION_BILINEAR);
                    graphics.setRenderingHint(
                            RenderingHints.KEY_ALPHA_INTERPOLATION, RenderingHints.VALUE_ALPHA_INTERPOLATION_QUALITY);
                    graphics.setRenderingHint(
                            RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                    graphics.setRenderingHint(
                            RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
                    graphics.setRenderingHint(
                            RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                    graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                            RenderingHints.VALUE_TEXT_ANTIALIAS_GASP);
                    // clear the drawing area
                    graphics.setPaint(Color.white);
                    graphics.fill(new Rectangle2D.Float(0, 0, dimensions.width, dimensions.height));

                    // render
                    slide.draw(graphics);
                    String outFile = FilenameUtils.concat(dirLocation.getAbsolutePath(),
                            slide.getSlideName() + ".jpeg");
                    FileOutputStream out = new FileOutputStream(outFile);
                    javax.imageio.ImageIO.write(img, "jpeg", out);
                    logger.info("Exported JPEG file: {}", outFile);
                }
            }
        }
    }

    private void importImagesToNewPPT(String outFileLocation, File dirLocation)
            throws IOException {
        try (FileOutputStream outFile = new FileOutputStream(outFileLocation)) {
            try (var ppt = new XMLSlideShow()) {
                Collection<File> files = FileUtils.listFiles(
                        dirLocation,
                        new RegexFileFilter("^(.*?)(.jpg|.png|.svg|.jpeg|.gif)"),
                        DirectoryFileFilter.DIRECTORY);

                // Files imported
                for (var file : files) {
                    var slide = ppt.createSlide();
                    byte[] pictureData = IOUtils.toByteArray(new FileInputStream(file));
                    XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                    var picture = slide.createPicture(pd);
                    var dimensions = slide.getSlideShow().getPageSize();
                    resizeToDimensions(picture, dimensions);
                    logger.info("Image imported: {}", file);
                }

                ppt.write(outFile);
                logger.info("Exported file: {}", outFileLocation);
            }
        }
    }
}

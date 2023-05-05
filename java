import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import javax.imageio.ImageIO;

public class PolyglotBMP {

    public static void main(String[] args) throws IOException {
        // Create a new BMP file.
        BufferedImage image = new BufferedImage(100, 100, BufferedImage.TYPE_INT_RGB);
        // Write the image to a file.
        ImageIO.write(image, "BMP", new File("polyglot.bmp"));
        // Create a new JAR file.
        File jarFile = new File("polyglot.jar");
        // Add the BMP file to the JAR file.
        new JarOutputStream(jarFile).putNextEntry("polyglot.bmp");
        ImageIO.write(image, "BMP", jarFile.toPath());
        // Add the hello world class to the JAR file.
        new JarOutputStream(jarFile).putNextEntry("helloworld.class");
        // Write the hello world class to the JAR file.
        byte[] helloWorldClass = helloWorld.class.getBytes();
        jarFile.toPath().write(helloWorldClass);
        // Run the hello world program.
        java.awt.Desktop.getDesktop().open(jarFile);
    }
}

class helloWorld {

    public static void main(String[] args) {
        System.out.println("Hello, world!");
    }
}

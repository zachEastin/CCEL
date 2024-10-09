# CCEL to Logos Personal Book
This is a personal python script to convert ThML into docx compatible with Logos Personal Book format

# Setup

# Install Java/Saxon-HE
*Saxon-HE is an open-source XSLT and XQuery processor developed in Java.*

### Step 1: Install Java
1. **Download Java Development Kit (JDK):**
   - Visit the [Oracle JDK download page](https://www.oracle.com/java/technologies/javase-jdk11-downloads.html) for a free version.
   - Download the installer for your Windows version.

2. **Install JDK:**
   - Run the downloaded installer and follow the installation instructions.
   - Make note of the installation path (usually something like `C:\Program Files\Java\jdk-11.x.x`).

3. **Set JAVA_HOME Environment Variable:**
   - Right-click on "This PC" or "My Computer" and select "Properties."
   - Click on "Advanced system settings" and then the "Environment Variables" button.
   - Under "System Variables," click "New" to create a new variable:
     - **Variable name:** `JAVA_HOME`
     - **Variable value:** The path where JDK is installed (e.g., `C:\Program Files\Java\jdk-11.x.x`).
   - Click "OK" to save the variable.

4. **Add Java to System PATH:**
   - In the "System Variables" section, find and select the `Path` variable, then click "Edit."
   - Click "New" and add the path to the Java `bin` directory, typically `C:\Program Files\Java\jdk-11.x.x\bin`.
   - Click "OK" to save and close all dialog boxes.

### Step 2: Download Saxon-HE
1. **Download Saxon-HE:**
   - Go to the [Saxon-HE download page](http://saxon.sourceforge.net/#Download).
   - Download the latest version of Saxon-HE (the `.zip` file).

2. **Extract the ZIP file:**
   - Right-click on the downloaded `.zip` file and select "Extract All..."
   - Set it to extract the files to `C:\SaxonHE` (create if necessary)

### Step 3: Run Saxon-HE
1. **Open Command Prompt:**
   - Press `Win + R`, type `cmd`, and press Enter.

2. **Navigate to Saxon Directory:**
   ```bash
   cd C:\SaxonHE\9.x.x.x
   ```
   Replace `9.x.x.x` with the actual version number of the extracted Saxon-HE folder.

3. **Run Saxon-HE:**
   You can run Saxon-HE using the following command:
   ```bash
   java -jar saxon-he-9.x.x.x.jar [options]
   ```
   Replace `9.x.x.x` with the actual version number.

### Step 4: Verify Installation
To verify that Saxon-HE is installed correctly, you can execute:
```bash
java -jar saxon-he-9.x.x.x.jar -s:input.xml -xsl:stylesheet.xsl -o:output.xml
```
Make sure you replace `input.xml`, `stylesheet.xsl`, and `output.xml` with your actual file names.

### Conclusion
You have now installed Saxon-HE on your Windows machine. You can use it for XSLT transformations or XQuery processing as needed. If you encounter any issues, make sure your Java installation is correct and that the environment variables are set properly.

# Install Pandoc

Install [Pandoc](https://pandoc.org/installing.html)
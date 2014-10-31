/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import junit.framework.Assert;
import junit.framework.TestCase;

import org.apache.commons.io.FileUtils;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class FileServiceImplTest extends TestCase
{
    String testDirPath = System.getenv("SIRECA_HOME") + "/test/";

    public FileServiceImplTest()
    {
        super();
    }

    @Before
    public void setUp() throws Exception
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        PrintWriter writer;

        // Create root test path
        fileService.addDirectory(testDirPath);

        // Create and fill folder for deleteFile
        fileService.addDirectory(testDirPath + "deleteFile");
        fileService.addFile(testDirPath + "deleteFile/file1.txt");

        // Create and fill folder for RemoveDirTest
        fileService.addDirectory(testDirPath + "deleteDir");
        fileService.addFile(testDirPath + "deleteDir/file1.txt");
        fileService.addFile(testDirPath + "deleteDir/file2.txt");
        fileService.addFile(testDirPath + "deleteDir/file3.txt");
        fileService.addFile(testDirPath + "deleteDir/file4.txt");

        // Create folder for AddFileTest
        fileService.addDirectory(testDirPath + "addFile");

        // Create and fill folder for getDirectory
        fileService.addDirectory(testDirPath + "getDir");
        fileService.addFile(testDirPath + "getDir/file1.txt");
        fileService.addFile(testDirPath + "getDir/file2.txt");
        fileService.addFile(testDirPath + "getDir/file3.txt");
        fileService.addFile(testDirPath + "getDir/file4.txt");

        // Create folder for getDateTest
        fileService.addDirectory(testDirPath + "getDate/");
        fileService.addFile(testDirPath + "getDate/file1.txt");

        // Create folder for getSizeTest and fill file
        fileService.addDirectory(testDirPath + "/getSize");

        writer = new PrintWriter(testDirPath + "getSize/testSize.txt", "UTF-8");
        writer.print("0123456789");
        writer.close();

        // Create folder for getExtensionTest
        fileService.addDirectory(testDirPath + "getExtension/");
        fileService.addFile(testDirPath + "getExtension/file1.ximi");

        // Create folder for GetContentTest and add a file with content
        fileService.addDirectory(testDirPath + "getContent");

        writer = new PrintWriter(testDirPath + "getContent/testContent.txt", "UTF-8");
        writer.print("This is a content test.");
        writer.close();

        // Create folder for FileCopyTest and add a file with content
        fileService.addDirectory(testDirPath + "fileCopy");

        writer = new PrintWriter(testDirPath + "fileCopy/testCopy.txt", "UTF-8");
        writer.println("Lorem ipsum dolor sit amet, consectetur adipiscing elit."
                + " Nunc et finibus massa. Quisque et tempus massa. Morbi vitae "
                + "odio luctus, viverra nisl in, fringilla nunc. Proin ut sapien a"
                + " erat suscipit sodales elementum non ipsum. Aenean scelerisque "
                + "dapibus nunc, eu hendrerit justo sagittis in. Proin quis sapien "
                + "in neque fringilla fringilla. Cras non sollicitudin dolor. "
                + "Proin eu iaculis tellus. Aliquam facilisis, nisi id volutpat "
                + "cursus, nibh turpis semper orci, vitae tristique magna purus "
                + "in dolor. Nam vestibulum, risus non dictum pharetra, dui urna "
                + "efficitur tellus, elementum pellentesque ligula ipsum vitae "
                + "justo. Nam vulputate lectus ac leo molestie, vitae fringilla"
                + " magna pellentesque. Nunc vitae scelerisque nisl.");
        writer.println("In lacus lacus, lobortis vel tincidunt ac, venenatis id "
                + "lacus. Suspendisse sit amet eleifend lorem, quis gravida justo. "
                + "Cras pellentesque ultricies urna a ullamcorper. Pellentesque"
                + " sollicitudin nibh a sagittis commodo. Ut in placerat leo. Aenean"
                + " elit ipsum, tempor in bibendum eu, molestie at lorem. Aliquam ex"
                + " diam, efficitur in quam eget, laoreet pellentesque ipsum. Nunc "
                + "pellentesque quis nunc ac tempus. Duis rutrum nulla ac felis maximus "
                + "placerat. Fusce eget quam neque. Maecenas felis nisl, eleifend id dolor"
                + " a, convallis malesuada ligula. Aenean blandit finibus justo, ac commodo"
                + " dolor. Vestibulum ante ipsum primis in faucibus orci.");
        writer.close();
    }

    @After
    public void tearDown() throws Exception
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        fileService.deleteDirectory(testDirPath);

    }

    @Test
    public void testContext()
    {
        Assert.assertNotNull(SpringApplicationContext.getBean("fileService"));
    }

    @Test
    public void testAddDirectory()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        assertTrue(fileService.addDirectory(testDirPath + "addDirectory"));
    }

    @Test
    public void testdeleteDirectory()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        assertTrue(fileService.deleteDirectory(testDirPath + "deleteDir"));
    }

    @Test
    public void testDeleteFile()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        assertTrue(fileService.deleteFile(testDirPath + "deleteFile/file1.txt"));
    }

    @Test
    public void testAddFile()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        assertTrue(fileService.addFile(testDirPath + "addFile/testAddFile.txt"));
    }

    @Test
    public void testGetDirectory()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        File[] file = fileService.getDirectory(testDirPath + "getDir");

        assertTrue(file[0].getName().equals("file1.txt"));
        assertTrue(file[1].getName().equals("file2.txt"));
        assertTrue(file[2].getName().equals("file3.txt"));
        assertTrue(file[3].getName().equals("file4.txt"));
    }

    @Test
    public void testGetFileDate()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        assertEquals(
                new SimpleDateFormat("dd-MM-yyyy").format(new Date()),
                new SimpleDateFormat("dd-MM-yyyy").format(fileService.getFileDate(testDirPath
                        + "getDate/file1.txt")));
    }

    @Test
    public void testGetFileSize()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        assertTrue(fileService.getFileSize(testDirPath + "getSize/testSize.txt") == 10);
    }

    @Test
    public void testGetFileExtension()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        File file = new File(testDirPath + "getExtension/file1.ximi");

        assertTrue(fileService.getFileExtension(file).equals("ximi"));
    }

    @Test
    public void testGetFileContent()
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        ArrayList<String> fileContent = null;

        try
        {
            fileContent = fileService.getFileContent(testDirPath
                    + "getContent/testContent.txt");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }

        assertTrue(fileContent.get(0).equals("This is a content test."));
    }

    @Test
    public void testFileCopy()
    {

        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        fileService.fileCopy(testDirPath + "fileCopy/testCopy.txt", testDirPath
                + "fileCopy/testCopyOut.txt");
        File file1 = new File(testDirPath + "fileCopy/testCopy.txt");
        File file2 = new File(testDirPath + "fileCopy/testCopyOut.txt");

        try
        {
            assertTrue(FileUtils.contentEquals(file1, file2));
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }

    }

}

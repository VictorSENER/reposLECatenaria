/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.IOException;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;

import junit.framework.Assert;
import junit.framework.TestCase;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.jacob.com.Variant;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class JACOBServiceImplTest extends TestCase
{

    public JACOBServiceImplTest()
    {
        super();
    }

    @Before
    public void setUp() throws Exception
    {
        IsJUnit.setJunitRunning(true);
    }

    @After
    public void tearDown() throws Exception
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        fileService.deleteFile("C:\\SIRECA/web/src/test/resources/testJunit - in.xlsx");
    }

    @Test
    public void testContext()
    {
        Assert.assertNotNull(SpringApplicationContext.getBean("jacobService"));
    }

    @Test
    public void testExecuteCoreCommand() throws NoSuchAlgorithmException,
            IOException
    {
        JACOBService jacobService = (JACOBService) SpringApplicationContext.getBean("jacobService");

        jacobService.executeCoreCommand("C:\\SIRECA/web/src/test/resources/",
                "", new ArrayList<Variant>());

        assertTrue(true);

    }

}

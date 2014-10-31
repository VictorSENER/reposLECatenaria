/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import junit.framework.Assert;
import junit.framework.TestCase;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.sener.sireca.web.bean.Catenaria;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class CatenariaServiceImplTest extends TestCase
{

    public CatenariaServiceImplTest()
    {
        super();
    }

    @Test
    public void testContext()
    {
        Assert.assertNotNull(SpringApplicationContext.getBean("catenariaService"));
    }

    @Test
    public void testGetAllCatenarias()
    {
        CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

        List<Catenaria> lcat = catenariaService.getAllCatenarias();
        assertTrue(lcat.get(0).getId() == 1);
        assertTrue(lcat.get(1).getId() == 2);
        assertTrue(lcat.get(2).getId() == 3);

        assertTrue(lcat.get(0).getNomCatenaria().equals("Test Catenaria 1"));
        assertTrue(lcat.get(1).getNomCatenaria().equals("Test Catenaria 2"));
        assertTrue(lcat.get(2).getNomCatenaria().equals("Test Catenaria 3"));
    }

    @Test
    public void testGetCatenariaById()
    {
        CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

        assertTrue(catenariaService.getCatenariaById(1).getNomCatenaria().equals(
                "Test Catenaria 1"));
        assertTrue(catenariaService.getCatenariaById(2).getNomCatenaria().equals(
                "Test Catenaria 2"));
        assertTrue(catenariaService.getCatenariaById(3).getNomCatenaria().equals(
                "Test Catenaria 3"));
    }

    @Test
    public void testGetCatenariaByTitle()
    {

        CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

        assertTrue(catenariaService.getCatenariaByTitle("Test Catenaria 1").getId() == 1);
        assertTrue(catenariaService.getCatenariaByTitle("Test Catenaria 2").getId() == 2);
        assertTrue(catenariaService.getCatenariaByTitle("Test Catenaria 3").getId() == 3);
    }

    @Test
    public void testGetListCatenarias()
    {
        CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

        List<String> lStr = catenariaService.getListCatenarias();

        assertTrue(lStr.get(0).equals("Test Catenaria 1"));
        assertTrue(lStr.get(1).equals("Test Catenaria 2"));
        assertTrue(lStr.get(2).equals("Test Catenaria 3"));
    }
}

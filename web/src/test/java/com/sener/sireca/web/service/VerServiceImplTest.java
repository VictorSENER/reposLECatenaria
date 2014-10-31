/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import static org.junit.Assert.assertTrue;

import java.util.List;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class VerServiceImplTest
{

    Project project;

    @Before
    public void setUp() throws Exception
    {
        IsJUnit.setJunitRunning(true);

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        int randomNum = 1 + (int) (Math.random() * 10000);

        String titulo = "Proyecto Test" + randomNum;
        String cliente = "Nombre Cliente";
        String referencia = "Referencia";

        project = new Project();
        project.setTitulo(titulo);
        project.setIdUsuario(1);
        project.setCliente(cliente);
        project.setReferencia(referencia);
        project.setIdCatenaria(1);

        // Store new project into DB.
        int id = projectService.insertProject(project);

        project = projectService.getProjectById(id); // Chapuza

        ReplanteoVersion repVer = replanteoService.getVersion(project, 1);

        for (int i = 0; i < 8; i++)
            replanteoService.createVersion(project);

        ReplanteoRevision repRev;
        for (int i = 0; i < 4; i++)
        {
            repRev = replanteoService.createRevision(repVer, 1, "Comment");
            fileService.addFile(repRev.getExcelPath());

        }

    }

    @After
    public void tearDown() throws Exception
    {
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

        projectService.deleteProject(project.getId());

    }

    @Test
    public void testContext()
    {

        Assert.assertNotNull(SpringApplicationContext.getBean("verService"));
    }

    @Test
    public void testGetVersions()
    {
        VerService verService = (VerService) SpringApplicationContext.getBean("verService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        List<Integer> verList = verService.getVersions(project.getCalcReplanteoBasePath());

        assertTrue(verList.size() == 9);

        for (int i = 0; i < verList.size(); i++)
            assertTrue(verList.get(i) == replanteoService.getVersion(project,
                    i + 1).getNumVersion());

    }

    @Test
    public void testGetVersion()
    {
        VerService verService = (VerService) SpringApplicationContext.getBean("verService");

        for (int i = 1; i < 9; i++)
            assertTrue(verService.getVersion(
                    project.getCalcReplanteoBasePath(), i));
    }

    @Test
    public void testGetLastVersion()
    {
        VerService verService = (VerService) SpringApplicationContext.getBean("verService");

        int lastVer = verService.getLastVersion(project.getCalcReplanteoBasePath());

        assertTrue(lastVer == 9);

    }
}

/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import static org.junit.Assert.assertNull;
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
public class ReplanteoServiceImplTest
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

        for (int i = 0; i < 4; i++)
            replanteoService.createVersion(project);

        ReplanteoRevision repRev;
        for (int i = 0; i < 4; i++)
        {
            repRev = replanteoService.createRevision(repVer, 0, "Comment");
            repRev.setCalculated(true);
            fileService.addFile(repRev.getExcelPath());
        }

    }

    @After
    public void tearDown()
    {

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

        projectService.deleteProject(project.getId());

        try
        {
            FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
            fileService.deleteDirectory(System.getenv("SIRECA_HOME")
                    + "/projectTest/");
        }
        catch (Exception ex)
        {
        }

    }

    @Test
    public void testContext()
    {
        Assert.assertNotNull(SpringApplicationContext.getBean("projectService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("replanteoService"));
    }

    @Test
    public void testGetVersion()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        ReplanteoVersion repVer = replanteoService.getVersion(project, 1);

        assertTrue(repVer.getNumVersion() == 1);
    }

    @Test
    public void testCreateVersion()
    {

        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        ReplanteoVersion repVer = replanteoService.createVersion(project);

        assertTrue(repVer.getNumVersion() == replanteoService.getLastVersion(project));

        replanteoService.deleteVersion(project, repVer.getNumVersion());

    }

    @Test
    public void testGetVersionList()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        List<Integer> versions = replanteoService.getVersionList(project);

        for (int i = 1; i < 6; i++)
            assertTrue(versions.get(i - 1) == i);
    }

    @Test
    public void testGetRevisionList()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        ReplanteoVersion repVer = replanteoService.getVersion(project, 1);

        List<Integer> revisions = replanteoService.getRevisionList(repVer);

        for (int i = 1; i < 5; i++)
            assertTrue(revisions.get(i - 1) == i);
    }

    @Test
    public void testGetVersions()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        List<ReplanteoVersion> listRepVer = replanteoService.getVersions(project);

        assertTrue(listRepVer.size() == 5);

        for (int i = 0; i < listRepVer.size(); i++)
            assertTrue(listRepVer.get(i).getNumVersion() == replanteoService.getVersion(
                    project, i + 1).getNumVersion());

    }

    @Test
    public void testGetLastRevision()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        ReplanteoVersion repVer = replanteoService.getVersion(project, 1);

        assertTrue(replanteoService.getLastRevision(repVer) == 4);

    }

    @Test
    public void testGetLastVersion()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        assertTrue(replanteoService.getLastVersion(project) == 5);

    }

    @Test
    public void testDeleteVersion()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        replanteoService.createVersion(project);

        replanteoService.deleteVersion(project, 6);

        ReplanteoVersion repVer = replanteoService.getVersion(project, 6);

        assertNull(repVer);

    }

    @Test
    public void testGetRevisions()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        ReplanteoVersion repVer = replanteoService.getVersion(project, 1);

        List<ReplanteoRevision> listRepRev = replanteoService.getRevisions(repVer);

        for (int i = 0; i < listRepRev.size(); i++)
            assertTrue(listRepRev.get(i).getNumRevision() == replanteoService.getRevision(
                    repVer, i + 1).getNumRevision());
    }

    @Test
    public void testGetRevision()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        ReplanteoVersion repVer = replanteoService.getVersion(project, 1);
        ReplanteoRevision repRev = replanteoService.getRevision(repVer, 1);

        assertTrue(repRev.getNumVersion() == 1);
        assertTrue(repRev.getNumRevision() == 1);
    }

    @Test
    public void testCreateRevision()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");

        fileService.addFile(repRev.getExcelPath());

        assertTrue(repRev.getNumRevision() == replanteoService.getLastRevision(repVer));
        try
        {
            replanteoService.deleteRevision(project, repRev.getNumVersion(),
                    repRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteFile(repRev.getExcelPath());

    }

    @Test
    public void testDeleteRevision()
    {
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");

        fileService.addFile(repRev.getExcelPath());

        try
        {
            replanteoService.deleteRevision(project, repRev.getNumVersion(),
                    repRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteFile(repRev.getExcelPath());

        assertNull(replanteoService.getRevision(repVer, 5));

    }

}

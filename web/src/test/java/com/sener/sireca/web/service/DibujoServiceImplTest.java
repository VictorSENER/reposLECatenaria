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

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class DibujoServiceImplTest
{

    Project project;

    @Before
    public void setUp() throws Exception
    {

        IsJUnit.setJunitRunning(true);

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
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
        project.setPendolado("HOLA");
        // project.setViaDoble(true);

        // Store new project into DB.
        int id = projectService.insertProject(project);

        project = projectService.getProjectById(id);

        DibujoVersion dibVer = dibujoService.getVersion(project, 1);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");
        fileService.addFile(repRev.getExcelPath());

        for (int i = 0; i < 4; i++)
            dibujoService.createVersion(project);

        DibujoRevision dibRev;

        for (int i = 0; i < 4; i++)
        {

            dibRev = dibujoService.createRevision(dibVer, repRev, "Comment");
            dibRev.setCalculated(true);
            fileService.addFile(dibRev.getAutoCadPath());

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
        Assert.assertNotNull(SpringApplicationContext.getBean("dibujoService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("replanteoService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("fileService"));
    }

    @Test
    public void testGetVersion()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        DibujoVersion dibVer = dibujoService.getVersion(project, 1);

        assertTrue(dibVer.getNumVersion() == 1);
    }

    @Test
    public void testCreateVersion()
    {

        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        DibujoVersion dibVer = dibujoService.createVersion(project);

        assertTrue(dibVer.getNumVersion() == dibujoService.getLastVersion(project));

        dibujoService.deleteVersion(project, dibVer.getNumVersion());

    }

    @Test
    public void testGetVersionList()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        List<Integer> versions = dibujoService.getVersionList(project);

        for (int i = 1; i < 6; i++)
            assertTrue(versions.get(i - 1) == i);
    }

    @Test
    public void testGetRevisionList()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        DibujoVersion dibVer = dibujoService.getVersion(project, 1);

        List<Integer> revisions = dibujoService.getRevisionList(dibVer);

        for (int i = 1; i < 5; i++)
            assertTrue(revisions.get(i - 1) == i);
    }

    @Test
    public void testGetVersions()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        List<DibujoVersion> listRepVer = dibujoService.getVersions(project);

        assertTrue(listRepVer.size() == 5);

        for (int i = 0; i < listRepVer.size(); i++)
            assertTrue(listRepVer.get(i).getNumVersion() == dibujoService.getVersion(
                    project, i + 1).getNumVersion());

    }

    @Test
    public void testGetLastRevision()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        DibujoVersion dibVer = dibujoService.getVersion(project, 1);

        assertTrue(dibujoService.getLastRevision(dibVer) == 4);

    }

    @Test
    public void testGetLastVersion()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        assertTrue(dibujoService.getLastVersion(project) == 5);

    }

    @Test
    public void testDeleteVersion()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        dibujoService.createVersion(project);

        dibujoService.deleteVersion(project, 6);

        DibujoVersion dibVer = dibujoService.getVersion(project, 6);

        assertNull(dibVer);

    }

    @Test
    public void testGetRevisions()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        DibujoVersion dibVer = dibujoService.getVersion(project, 1);

        List<DibujoRevision> listRepRev = dibujoService.getRevisions(dibVer);

        for (int i = 0; i < listRepRev.size(); i++)
            assertTrue(listRepRev.get(i).getNumRevision() == dibujoService.getRevision(
                    dibVer, i + 1).getNumRevision());
    }

    @Test
    public void testGetRevision()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");

        DibujoVersion dibVer = dibujoService.getVersion(project, 1);
        DibujoRevision dibRev = dibujoService.getRevision(dibVer, 1);

        assertTrue(dibRev.getNumVersion() == 1);
        assertTrue(dibRev.getNumRevision() == 1);
    }

    @Test
    public void testCreateRevision()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        DibujoVersion dibVer = dibujoService.createVersion(project);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");
        fileService.addFile(repRev.getExcelPath());

        DibujoRevision dibRev = dibujoService.createRevision(dibVer, repRev,
                "Comment");

        fileService.addFile(dibRev.getAutoCadPath());

        assertTrue(dibRev.getNumRevision() == dibujoService.getLastRevision(dibVer));
        try
        {
            dibujoService.deleteRevision(project, dibRev.getNumVersion(),
                    dibRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteFile(dibRev.getAutoCadPath());

    }

    @Test
    public void testDeleteRevision()
    {
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        DibujoVersion dibVer = dibujoService.createVersion(project);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");
        DibujoRevision dibRev = dibujoService.createRevision(dibVer, repRev,
                "Comment");

        fileService.addFile(dibRev.getAutoCadPath());

        try
        {
            dibujoService.deleteRevision(project, dibRev.getNumVersion(),
                    dibRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteFile(dibRev.getAutoCadPath());

        assertNull(dibujoService.getRevision(dibVer, 5));

    }

}

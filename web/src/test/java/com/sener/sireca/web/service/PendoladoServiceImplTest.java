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

import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class PendoladoServiceImplTest
{

    Project project;

    @Before
    public void setUp() throws Exception
    {

        IsJUnit.setJunitRunning(true);

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
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

        project = projectService.getProjectById(id);

        PendoladoVersion penVer = pendoladoService.getVersion(project, 1);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");
        fileService.addFile(repRev.getExcelPath());

        for (int i = 0; i < 4; i++)
            pendoladoService.createVersion(project);

        PendoladoRevision penRev;

        for (int i = 0; i < 4; i++)
        {
            System.out.println("|-------> Crea la revisión " + (i + 1));
            penRev = pendoladoService.createRevision(penVer, repRev, "Comment");
            penRev.setCalculated(true);
            fileService.addFile(penRev.getPDFPath());
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
        Assert.assertNotNull(SpringApplicationContext.getBean("pendoladoService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("replanteoService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("fileService"));
    }

    @Test
    public void testGetVersion()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        PendoladoVersion penVer = pendoladoService.getVersion(project, 1);

        assertTrue(penVer.getNumVersion() == 1);
    }

    @Test
    public void testCreateVersion()
    {

        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        PendoladoVersion penVer = pendoladoService.createVersion(project);

        assertTrue(penVer.getNumVersion() == pendoladoService.getLastVersion(project));

        pendoladoService.deleteVersion(project, penVer.getNumVersion());

    }

    @Test
    public void testGetVersionList()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        List<Integer> versions = pendoladoService.getVersionList(project);

        for (int i = 1; i < 6; i++)
            assertTrue(versions.get(i - 1) == i);
    }

    @Test
    public void testGetRevisionList()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        PendoladoVersion penVer = pendoladoService.getVersion(project, 1);

        List<Integer> revisions = pendoladoService.getRevisionList(penVer);

        for (int i = 1; i < 5; i++)
            assertTrue(revisions.get(i - 1) == i);
    }

    @Test
    public void testGetVersions()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        List<PendoladoVersion> listRepVer = pendoladoService.getVersions(project);

        assertTrue(listRepVer.size() == 5);

        for (int i = 0; i < listRepVer.size(); i++)
            assertTrue(listRepVer.get(i).getNumVersion() == pendoladoService.getVersion(
                    project, i + 1).getNumVersion());

    }

    @Test
    public void testGetLastRevision()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        PendoladoVersion penVer = pendoladoService.getVersion(project, 1);

        assertTrue(pendoladoService.getLastRevision(penVer) == 4);

    }

    @Test
    public void testGetLastVersion()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        assertTrue(pendoladoService.getLastVersion(project) == 5);

    }

    @Test
    public void testDeleteVersion()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        pendoladoService.createVersion(project);

        pendoladoService.deleteVersion(project, 6);

        PendoladoVersion penVer = pendoladoService.getVersion(project, 6);

        assertNull(penVer);

    }

    @Test
    public void testGetRevisions()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        PendoladoVersion penVer = pendoladoService.getVersion(project, 1);

        List<PendoladoRevision> listRepRev = pendoladoService.getRevisions(penVer);

        for (int i = 0; i < listRepRev.size(); i++)
            assertTrue(listRepRev.get(i).getNumRevision() == pendoladoService.getRevision(
                    penVer, i + 1).getNumRevision());
    }

    @Test
    public void testGetRevision()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");

        PendoladoVersion penVer = pendoladoService.getVersion(project, 1);
        PendoladoRevision penRev = pendoladoService.getRevision(penVer, 1);

        assertTrue(penRev.getNumVersion() == 1);
        assertTrue(penRev.getNumRevision() == 1);
    }

    @Test
    public void testCreateRevision()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        PendoladoVersion penVer = pendoladoService.createVersion(project);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");
        fileService.addFile(repRev.getExcelPath());

        PendoladoRevision penRev = pendoladoService.createRevision(penVer,
                repRev, "Comment");

        fileService.addFile(penRev.getPDFPath());

        assertTrue(penRev.getNumRevision() == pendoladoService.getLastRevision(penVer));
        try
        {
            pendoladoService.deleteRevision(project, penRev.getNumVersion(),
                    penRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteFile(penRev.getPDFPath());

    }

    @Test
    public void testDeleteRevision()
    {
        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        PendoladoVersion penVer = pendoladoService.createVersion(project);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "Comment");
        PendoladoRevision penRev = pendoladoService.createRevision(penVer,
                repRev, "Comment");

        fileService.addFile(penRev.getPDFPath());

        try
        {
            pendoladoService.deleteRevision(project, penRev.getNumVersion(),
                    penRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteFile(penRev.getPDFPath());

        assertNull(pendoladoService.getRevision(penVer, 5));

    }

}

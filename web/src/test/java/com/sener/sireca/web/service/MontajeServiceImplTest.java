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

import com.sener.sireca.web.bean.MontajeRevision;
import com.sener.sireca.web.bean.MontajeVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class MontajeServiceImplTest
{

    Project project;

    @Before
    public void setUp() throws Exception
    {

        IsJUnit.setJunitRunning(true);

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
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

        MontajeVersion monVer = montajeService.getVersion(project, 1);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "");
        fileService.addFile(repRev.getExcelPath());

        for (int i = 0; i < 4; i++)
            montajeService.createVersion(project);

        MontajeRevision monRev;

        for (int i = 0; i < 4; i++)
        {
            monRev = montajeService.createRevision(monVer, repRev, "");
            monRev.setCalculated(true);
            fileService.addDirectory(monRev.getBasePath());
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
        Assert.assertNotNull(SpringApplicationContext.getBean("montajeService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("replanteoService"));
        Assert.assertNotNull(SpringApplicationContext.getBean("fileService"));
    }

    @Test
    public void testGetVersion()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        MontajeVersion monVer = montajeService.getVersion(project, 1);

        assertTrue(monVer.getNumVersion() == 1);
    }

    @Test
    public void testCreateVersion()
    {

        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        MontajeVersion monVer = montajeService.createVersion(project);

        assertTrue(monVer.getNumVersion() == montajeService.getLastVersion(project));

        montajeService.deleteVersion(project, monVer.getNumVersion());

    }

    @Test
    public void testGetVersionList()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        List<Integer> versions = montajeService.getVersionList(project);

        for (int i = 1; i < 6; i++)
            assertTrue(versions.get(i - 1) == i);
    }

    @Test
    public void testGetRevisionList()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        MontajeVersion monVer = montajeService.getVersion(project, 1);

        List<Integer> revisions = montajeService.getRevisionList(monVer);

        for (int i = 1; i < 5; i++)
            assertTrue(revisions.get(i - 1) == i);
    }

    @Test
    public void testGetVersions()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        List<MontajeVersion> listRepVer = montajeService.getVersions(project);

        assertTrue(listRepVer.size() == 5);

        for (int i = 0; i < listRepVer.size(); i++)
            assertTrue(listRepVer.get(i).getNumVersion() == montajeService.getVersion(
                    project, i + 1).getNumVersion());

    }

    @Test
    public void testGetLastRevision()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        MontajeVersion monVer = montajeService.getVersion(project, 1);

        assertTrue(montajeService.getLastRevision(monVer) == 4);

    }

    @Test
    public void testGetLastVersion()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        assertTrue(montajeService.getLastVersion(project) == 5);

    }

    @Test
    public void testDeleteVersion()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        montajeService.createVersion(project);

        montajeService.deleteVersion(project, 6);

        MontajeVersion monVer = montajeService.getVersion(project, 6);

        assertNull(monVer);

    }

    @Test
    public void testGetRevisions()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        MontajeVersion monVer = montajeService.getVersion(project, 1);

        List<MontajeRevision> listRepRev = montajeService.getRevisions(monVer);

        for (int i = 0; i < listRepRev.size(); i++)
            assertTrue(listRepRev.get(i).getNumRevision() == montajeService.getRevision(
                    monVer, i + 1).getNumRevision());
    }

    @Test
    public void testGetRevision()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");

        MontajeVersion monVer = montajeService.getVersion(project, 1);
        MontajeRevision monRev = montajeService.getRevision(monVer, 1);

        assertTrue(monRev.getNumVersion() == 1);
        assertTrue(monRev.getNumRevision() == 1);
    }

    @Test
    public void testCreateRevision()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        MontajeVersion monVer = montajeService.createVersion(project);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "");
        fileService.addFile(repRev.getExcelPath());

        MontajeRevision monRev = montajeService.createRevision(monVer, repRev,
                "");

        fileService.addDirectory(monRev.getBasePath());

        assertTrue(monRev.getNumRevision() == montajeService.getLastRevision(monVer));
        try
        {
            montajeService.deleteRevision(project, monRev.getNumVersion(),
                    monRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteDirectory(monRev.getBasePath());

    }

    @Test
    public void testDeleteRevision()
    {
        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");

        MontajeVersion monVer = montajeService.createVersion(project);
        ReplanteoVersion repVer = replanteoService.createVersion(project);
        ReplanteoRevision repRev = replanteoService.createRevision(repVer, 1,
                "");
        MontajeRevision monRev = montajeService.createRevision(monVer, repRev,
                "");

        fileService.addDirectory(monRev.getBasePath());

        try
        {
            montajeService.deleteRevision(project, monRev.getNumVersion(),
                    monRev.getNumRevision());
        }
        catch (Exception ex)
        {
        }

        fileService.deleteDirectory(monRev.getBasePath());

        assertNull(montajeService.getRevision(monVer, 5));

    }

}

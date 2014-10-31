/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import junit.framework.Assert;
import junit.framework.TestCase;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class ProjectServiceImplTest extends TestCase
{

    int id;
    String titulo;
    String cliente;
    String referencia;

    public ProjectServiceImplTest()
    {
        super();
    }

    @Override
    @Before
    public void setUp() throws Exception
    {

        IsJUnit.setJunitRunning(true);
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

        int randomNum = 1 + (int) (Math.random() * 10000);

        titulo = "Proyecto Test" + randomNum;
        cliente = "Nombre Cliente";
        referencia = "Referencia";

        Project project = new Project();
        project.setTitulo(titulo);
        project.setIdUsuario(1);
        project.setCliente(cliente);
        project.setReferencia(referencia);
        project.setIdCatenaria(1);

        // Store new project into DB.
        id = projectService.insertProject(project);

    }

    @Override
    @After
    public void tearDown() throws Exception
    {

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

        projectService.deleteProject(id);

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
    }

    @Test
    public void testGetProjectById()
    {

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        Project project = projectService.getProjectById(id);

        assertTrue(project.getTitulo().equals(titulo));
        assertTrue(project.getCliente().equals(cliente));
        assertTrue(project.getReferencia().equals(referencia));
    }

    @Test
    public void testGetProjectByTitle()
    {

        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        Project project = projectService.getProjectByTitle(titulo);

        assertTrue(project.getId() == id);
        assertTrue(project.getCliente().equals(cliente));
        assertTrue(project.getReferencia().equals(referencia));
    }

    @Test
    public void testUpdateProject()
    {
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

        titulo += " Edited";
        cliente += " Edited";
        referencia += " Edited";

        Project project = projectService.getProjectById(id);

        project.setTitulo(titulo);
        project.setCliente(cliente);
        project.setReferencia(referencia);

        try
        {
            projectService.updateProject(project);
        }
        catch (Exception e)
        {

        }
        project = projectService.getProjectById(id);

        assertTrue(project.getTitulo().equals(titulo));
        assertTrue(project.getCliente().equals(cliente));
        assertTrue(project.getReferencia().equals(referencia));
    }

}

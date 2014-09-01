/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import com.sener.sireca.web.bean.Project;

public interface ProjectService
{
    public int insertProject(Project project);

    public List<Project> getAllProjects();

    public Project getProjectById(int id);

    public Project getProjectByTitle(String title);

    public int updateProject(Project project);

    public int deleteProject(int id);

}

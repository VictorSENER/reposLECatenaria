/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.dao;

import java.util.List;

import com.sener.sireca.web.bean.Project;

public interface ProjectDao
{
    public int insertProject(Project project);

    public List<Project> getAllProjects();

    public Project getProjectById(int id);

    public int updateProject(Project project);

    public int deleteProject(int id);
}
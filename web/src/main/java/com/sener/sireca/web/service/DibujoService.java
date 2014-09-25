/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;

public interface DibujoService
{
    public List<DibujoVersion> getVersions(Project project);

    public DibujoVersion getVersion(Project project, int numVersion);

    public DibujoVersion createVersion(Project project);

    public void deleteVersion(Project project, int numVersion);

    public List<DibujoRevision> getRevisions(DibujoVersion version);

    public DibujoRevision getRevision(DibujoVersion version, int numRevision);

    public DibujoRevision createRevision(DibujoVersion version, int type);

    public void calculateRevision(DibujoRevision revision);

    public void deleteRevision(Project project, int numVersion, int numRevision);

    public int getLastVersion(Project project);

    public List<Integer> getRevisionList(DibujoVersion version);

    public List<Integer> getVersionList(Project project);
}

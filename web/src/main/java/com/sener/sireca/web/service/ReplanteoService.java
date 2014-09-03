/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;

public interface ReplanteoService
{
    public List<ReplanteoVersion> getVersions(Project project);

    public ReplanteoVersion getVersion(Project project, int numVersion);

    public ReplanteoVersion createVersion(Project project);

    public void deleteVersion(Project project, int numVersion);

    public List<ReplanteoRevision> getRevisions(ReplanteoVersion version);

    public ReplanteoRevision getRevision(ReplanteoVersion version,
            int numRevision);

    public ReplanteoRevision createRevision(ReplanteoVersion version, int type);

    public void calculateRevision(ReplanteoRevision revision);

    public void deleteRevision(Project project, int numVersion, int numRevision);

}

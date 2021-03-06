/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;

public interface ReplanteoService
{
    public List<ReplanteoVersion> getVersions(Project project);

    public ReplanteoVersion getVersion(Project project, int numVersion);

    public List<Integer> getVersionList(Project project);

    public ReplanteoVersion createVersion(Project project);

    public void deleteVersion(Project project, int numVersion);

    public List<ReplanteoRevision> getRevisions(ReplanteoVersion version);

    public List<Integer> getRevisionList(ReplanteoVersion version);

    public int getLastRevision(ReplanteoVersion version);

    public ReplanteoRevision getRevision(ReplanteoVersion version,
            int numRevision);

    public ReplanteoRevision createRevision(ReplanteoVersion version, int type);

    public void calculateRevision(ReplanteoRevision revision, long pkIni,
            long pkFin, String catenaria);

    public boolean deleteRevision(Project project, int numVersion,
            int numRevision);

    public String[] getProgressInfo(ReplanteoRevision revision)
            throws IOException;

    public int getLastVersion(Project project);

    public ArrayList<String> getErrorLog(ReplanteoRevision replanteoRevision)
            throws IOException;
}

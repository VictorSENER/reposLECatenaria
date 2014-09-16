/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import javax.servlet.http.HttpSession;

public interface ActiveProjectService
{

    public void setActive(HttpSession session, int idProj, String titleProj);

    public int getIdActive(HttpSession session);

}

/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.controller;

import java.util.Map;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Page;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.util.Initiator;

import com.sener.sireca.web.service.AuthenticationService;
import com.sener.sireca.web.session.UserCredential;
import com.sener.sireca.web.util.SpringApplicationContext;

public class AuthenticationInit implements Initiator
{

    public void doInit(Page page, Map<String, Object> args) throws Exception
    {
        HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();
        AuthenticationService authService = (AuthenticationService) SpringApplicationContext.getBean("authService");
        UserCredential cre = authService.getUserCredential(session);
        if (cre == null)
        {
            Executions.sendRedirect("/login");
            return;
        }
    }
}
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

import com.sener.sireca.web.bean.User;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = { "/applicationContext-servlet.xml" })
public class UserServiceImplTest extends TestCase
{

    int id;
    String username;
    String password;

    public UserServiceImplTest()
    {
        super();
    }

    @Override
    @Before
    public void setUp() throws Exception
    {

        IsJUnit.setJunitRunning(true);
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");

        int randomNum = 1 + (int) (Math.random() * 10000);

        username = "Username Test " + randomNum;
        password = "Password Test";

        User user = new User();
        user.setUsername(username);
        user.setPassword(password);

        // Store new project into DB.
        id = userService.insertUser(user);

    }

    @Override
    @After
    public void tearDown() throws Exception
    {
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");

        userService.deleteUser(id);
    }

    @Test
    public void testContext()
    {

        Assert.assertNotNull(SpringApplicationContext.getBean("userService"));
    }

    @Test
    public void testGetUserById()
    {

        UserService userService = (UserService) SpringApplicationContext.getBean("userService");
        User user = userService.getUserById(id);

        assertTrue(user.getPassword().equals(password));
        assertTrue(user.getUsername().equals(username));
    }

    @Test
    public void testGetProjectByTitle()
    {

        UserService userService = (UserService) SpringApplicationContext.getBean("userService");
        User user = userService.getUserByUsername(username);

        assertTrue(user.getId() == id);
        assertTrue(user.getPassword().equals(password));
    }

    @Test
    public void testUpdateProject()
    {
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");

        username += " Edited";
        password += " Edited";

        User user = userService.getUserById(id);
        user.setUsername(username);
        user.setPassword(password);

        try
        {
            userService.updateUser(user);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

        user = userService.getUserById(id);

        assertTrue(user.getPassword().equals(password));
        assertTrue(user.getUsername().equals(username));

    }

}

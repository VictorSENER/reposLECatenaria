/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("jacobService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class JACOBServiceImpl implements JACOBService

{
    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

    private List<File> prepareFiles(String path, String fase)
    {
        String initPath = System.getenv("SIRECA_HOME") + "/core/" + fase
                + ".xlsm";
        String finalPath = path + ".xlsm";

        fileService.fileCopy(initPath, finalPath);

        List<File> files = new ArrayList<File>();

        files.add(new File(path + ".xlsm"));
        files.add(new File(path + ".xlsx"));

        return files;

    }

    private void createCall(ActiveXComponent excel, String excelName,
            List<Variant> parameters)
    {
        int nParam = parameters.size();

        // Call the macro
        switch (nParam)
        {
            case 0:
                Dispatch.call(excel, "Run", new Variant(excelName
                        + "!ExecuteExcel"));

            case 1:
                Dispatch.call(excel, "Run", new Variant(excelName
                        + "!ExecuteExcel"), parameters.get(0));
                break;

            case 2:
                Dispatch.call(excel, "Run", new Variant(excelName
                        + "!ExecuteExcel"), parameters.get(0),
                        parameters.get(1));
                break;

            case 3:
                Dispatch.call(excel, "Run", new Variant(excelName
                        + "!ExecuteExcel"), parameters.get(0),
                        parameters.get(1), parameters.get(2));
                break;

            case 4:
                Dispatch.call(excel, "Run", new Variant(excelName
                        + "!ExecuteExcel"), parameters.get(0),
                        parameters.get(1), parameters.get(2), parameters.get(3));
                break;

        }

    }

    private void executeMacro(ActiveXComponent excel, List<File> files,
            List<Variant> parameters)
    {

        excel.setProperty("Visible", new Variant(true));

        final Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
        final Dispatch workBookConMacro = Dispatch.call(workbooks, "Open",
                files.get(0).getAbsolutePath()).toDispatch();
        final Dispatch workBookATratar = Dispatch.call(workbooks, "Open",
                files.get(1).getAbsolutePath()).toDispatch();

        createCall(excel, files.get(0).getName(), parameters);

        // Save and Close
        Dispatch.call(workBookATratar, "Save");
        // Close
        Dispatch.call(workBookConMacro, "Close", 0);
    }

    @Override
    public boolean executeCoreCommand(String path, String fase,
            List<Variant> parameters)
    {
        boolean todoOk = true;
        List<File> files = null;

        ComThread.InitSTA();

        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");

        try
        {
            files = prepareFiles(path, fase);

            executeMacro(excel, files, parameters);

        }
        catch (Exception e)
        {
            todoOk = false;
        }
        finally
        {
            try
            {
                excel.invoke("Quit", new Variant[0]);
                ComThread.Release();

                fileService.deleteFile(files.get(0).getAbsolutePath());

            }
            catch (Exception e)
            {
                // Do Nothing
            }
        }

        return todoOk;
    }
}

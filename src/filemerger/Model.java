/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package filemerger;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

/**
 *
 * @author mmercado
 */
public class Model {

    private HashMap<String, String> passwords;
    private File[] files;
    private ArrayList<Form> forms;

    public ArrayList<Form> getForms() {
        return forms;
    }

    public void setForms(ArrayList<Form> forms) {
        this.forms = forms;
    }
        
    public File[] getFiles() {
        return files;
    }

    public void setFiles(File[] files) {
        this.files = files;
    }

    public Model() {
        passwords = new HashMap();
    }

    public HashMap<String, String> getPasswords() {
        return passwords;
    }

    public void setPasswords(HashMap<String, String> passwords) {
        this.passwords = passwords;
    }

}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package filemerger;


/**
 *
 * @author mmercado
 */
public class FileMerger {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        //Model m = new Model("Sylvain", "Saurel");
        View v = new View("File Merger");
        Model m = new Model();
        Controller c = new Controller(m,v);
        c.initController();
        
    }

}

package docxtemplateexpander;

import java.util.List;

import com.google.common.collect.ImmutableList;
import com.google.common.collect.Lists;

public final class Substitution {
    private final String value;
//    private final List<String> notes = Lists.newArrayList();

    public Substitution(String value) {
        this.value = value;
    }
    
//    public Substitution addNote(String note) {
//        synchronized (notes) {
//            notes.add(note);
//        }
//        
//        return this;
//    }
    
    String getValue() {
        return value;
    }
    
//    ImmutableList<String> getNotes() {
//        return ImmutableList.copyOf(notes);
//    }
}

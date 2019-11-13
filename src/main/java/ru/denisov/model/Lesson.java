package ru.denisov.model;

import lombok.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Getter
@Setter
@ToString
@NoArgsConstructor
public
class Lesson {
    private static final Logger log = LoggerFactory.getLogger(Lesson.class);
    private String subject;
    private String type;
    private String teacher;
    private String auditory;
    private int lessonNumber;
}

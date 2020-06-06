# Excel-in-space
Excel In Space: An RPG Adventure Game built in Microsoft Excel for fun


# What Exists Currently
Excel In Space as it currently exists features:

    * Basic XML Reading / Writing Schema
    * Basic Cross Workbook Functionality
    * Planet Generation / Traversal on foot
    * Star System Generation / Traversal in ship
    * Inventory System


# To Do
I would like to formalize and expand these systems:

    * XML Reading / Writing Schema
        * Refactor and make more efficient

    * Cross Workbook Functionality
        * Establish 'Menu' Workbook to navigate between:
            * Body
            * Subtle Body
            * Brain
            * Compactor
        * 'Body' in this case would be either:
            * On Ship
            * On Planet
        * Remove any navigation to Character / Compactor from Bridge / planetWorkbook

    * Character System
        * This system will be like a mech-building mechanic, where you build out your robot body with all sorts of parts that affect all aspects of your adventuring. 
            * Define Body part list
                * Part types / stats
            * Define how parts interact with combat
                * Different classes of combat parts work differently
                * Different arrangements of parts offer different protections
            * Define How parts interact with adventuring
                * Different classes of parts offer different perks while adventuring

    * Subtle Body System
        * This system will look like a clubhouse sim game, where the better your clubhouse is, the more folk will want to come hang out in it. 
        * You create your clubhouse through knitting with Yaern
            * Hue / Saturation / Lightness of yaern used to knit items in the clubhouse affect which NPCs / spirits are attracted to clubhouse
        * NPCs you meet can come visit the clubhouse, which in turn unlocks buffs to apply to character
            * NPCs colors determine buff type
            * Fray determines how effective buffs are

    * Color System
        * Hue affects personality type
        * Saturation affects value of things / extremity of personality type
        * Lightness affects... something

    * NPC System
        * The Planets in the game should have plenty of characters to interact with
            * NPCs come in specific colors (Hue / Saturation / Lightness)
            * Meeting an NPC allows them to find your Subtle Body clubhouse
            * NPCs also can be interacted with in the following ways:
                * Conversation
                * Combat
                * Commerce

    * Combat System
        * The universe is a dangerous place, and as such, there exist plenty of things to fight!
            * I don't want combat that exists to whittle down opposing health bars
            * I do want combat that is part specific to both the attacker and defender

    * Conversation system
        * NPCs are fun to talk to! At least, I would like them to be.
            * Stock conversation material
                * Ask about Local Rumors
                * Ask about personal matters
                * Ask about political matters

    * Compactor (Inventory)
        * There will be two parts to one's inventory:
            * Physical Stuff
                * Materials, Parts, etc.
            * Subtle Stuff
                * Smoerph, Spoermph, Yaern, Information, etc.

    * Brain Tank
        * There will be a network-like system where one stores knowledge gained throughout playing
            * Similarly domained knowledge when networked together produces multiplicative results
            * Only have so much space in the brain tank to put knowledge
            * Knowledge accumulates whenever one does almost anything

    * Menu
        * Allows navigation between the workbooks used to instantiate all these systems
        * Has a log where notifications are put
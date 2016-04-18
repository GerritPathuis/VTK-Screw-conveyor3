Imports System.Math
Imports System
Imports System.Globalization
Imports System.Threading


Public Class Form1
    'Materials name; CEMA Material code; Conveyor loading; Component group, density min, Density max, HP Material
    Public Shared inputs() As String = {"Adipic Acid;45A35;30A;2B;720;720;0.5",
"Alfalfa Meal;18B45WY;30A;2D;220;350;0.6",
"Alfalfa Pellets;42C25;45;2D;660;690;0.5",
"Alfalfa Seed;13B15N;45;1A,1B,1C;160;240;0.4",
"Almonds, Broken;29C35Q;30A;2D;430;480;0.9",
"Almonds, Whole Shelled;29C35Q;30A;2D;450;480;0.9",
"Alum, Fine;48B35U;30A;3D;720;800;0.6",
"Alum, Lumps;55B25;45;2A,2B;800;960;1.4",
"Alumina;58B27MY;15;3D;880;1040;1.8",
"Alumina Fines;35A27MY;15;3D;560;560;1.6",
"Alumina Sized or Briquette;65D37;15;3D;1040;1040;2",
"Aluminate Gel (Aluminate Hydroxide);45B35;30B;2D;720;720;1.7",
"Aluminum Chips, Dry;11E45V;30A;2D;110;240;1.2",
"Aluminum Chips, Oily;11E45VY;30A;2D;110;240;0.8",
"Aluminum Hydrate;17C35;30A;1A,1B,1C;210;320;1.4",
"Aluminum Oxide;90A17MN;15;3D;960;1920;1.8",
"Aluminum Silicate (Andalusite);49C35S;45;3A,3B;780;780;0.7",
"Aluminum Sulfate;52C25;45;1A,1B,1C;720;930;1.3",
"Ammonium Chloride, Crystalline;49A45FRS;30A;1A,1B,1C;720;830;1",
"Ammonium Nitrate;54A35NTU;30A;3D;720;990;1.6",
"Ammonium Sulfate;52C35FOTU;30A;1A,1B,1C;720;930;1",
"Apple Pomace, Dry;15C45Y;30B;2D;240;240;1",
"Arsenate of Lead (Lead Arsenate);72A35R;30A;1A,1B,1C;1150;1150;1.4",
"Arsenic Pulverized;30A25R;45;2D;480;480;1",
"Asbestos-Rock (Ore);81D37R;15;3D;1300;1300;2",
"Asbestos-Shredded;30E46XY;30B;2D;320;640;1",
"Ash, Black Ground;105B35;30A;1A,1B,1C;1680;1680;2.5",
"Ashes, Coal, dry, 1⁄2 inch;40C46TY;30B;3D;560;720;3",
"Ashes, Coal, dry, 3 inch;38D46T;30B;3D;560;640;2.5",
"Ashes, Coal, Wet, 1⁄2 inch;48C46T;30B;3D;720;800;3",
"Ashes, Coal, Wet, 3 inch;48D46T;30B;3D;720;800;4",
"Ashes, Fly (Fly Ash);38A36M;30B;3D;480;720;2",
"Aspartic Acid;42A35XPLO;30A;1A,1B,1C;530;820;1.5",
"Asphalt, Crushed, 1⁄2 inch;45C45;30A;1A,1B,1C;720;720;2",
"Bagasse;9E45RVXY;30A;2A,2B,2C;110;160;1.5",
"Bakelite, Fine;38B25;45;1A,1B,1C;480;720;1.4",
"Baking Powder;48A35;30A;1B;640;880;0.6",
"Baking Soda (Sodium Bicarbonate);48A25;45;1B;640;880;0.6",
"Barite (Barium Sulfate), 1⁄2 to 3 inch;150D36;30B;3D;1920;2880;2.6",
"Barite, Powder;150A35X;30A;2D;1920;2880;2",
"Barium Carbonate;72A45R;30A;2D;1150;1150;1.6",
"Bark, Wood, Refuse;15E45TVY;30A;3D;160;320;2",
"Barley, Fine, Ground;31B35;30A;1A,1B,1C;380;610;0.4",
"Barley, Malted;31C35;30A;1A,1B,1C;500;500;0.4",
"Barley, Meal;28C35;30A;1A,1B,1C;450;450;0.4",
"Barley, Whole;42B25N;45;1A,1B,1C;580;770;0.5",
"Basalt;93B27;15;3D;1280;1680;1.8",
"Bauxite, Crushed, 3 inch (Aluminum Ore);80D36;30B;3D;1200;1360;2.5",
"Bauxite, Dry, Ground(Aluminum Ore);68B25;45;2D;1090;1090;1.8",
"Beans, Castor, Meal;38B35W;30A;1A,1B,1C;560;640;0.8",
"Beans, Castor, Whole Shelled;36C15W;45;1A,1B,1C;580;580;0.5",
"Beans, Navy, Dry;48C15;45;1A,1B,1C;770;770;0.5",
"Beans, Navy, Steeped;60C25;45;1A,1B,1C;960;960;0.8",
"Bentonite, 100 Mesh;55A25MXY;45;2D;800;960;0.7",
"Bentonite, Crude;37D45X;30A;2D;540;640;1.2",
"Benzene Hexachloride;56A45R;30A;1A,1B,1C;900;900;0.6",
"Bicarbonate of Soda (Baking Soda);48A25;45;1B;640;880;0.6",
"Blood, Dried;40D45U;30A;2D;560;720;2",
"Blood, Ground, Dried;30A35U;30A;1A,1B;480;480;1",
"Bone Ash (Tricalcium Phosphate);45A45;30A;1A,1B;640;800;1.6",
"Boneblack;23A25Y;45;1A,1B;320;400;1.5",
"Bonechar;34B35;30A;1A,1B;430;640;1.6",
"Bonemeal;55B35;30A;2D;800;960;1.7",
"Bones, Crushed;43D45;30A;2D;560;800;2",
"Bones, Ground;50B35;30A;2D;800;800;1.7",
"Bones, Whole**;43E45V;30A;2D;560;800;3",
"Borate of Lime;60A35;30A;1A,1B,1C;960;960;0.6",
"Borax Screening, 1⁄2 inch;58C35;30A;2D;880;960;1.5",
"Borax, 1-1⁄2  to 2 inch Lump;58D35;30A;2D;880;960;1.8",
"Borax, 2 to 3 inch Lump;65D35;30A;2D;960;1120;2",
"Borax, Fine;50B25T;45;3D;720;880;0.7",
"Boric Acid, Fine;55B25T;45;3D;880;880;0.8",
"Boron;75A37;15;2D;1200;1200;1",
"Bran, Rice-Rye-Wheat;18B355NY;30A;1A,1B,1C;260;320;0.5",
"Braunite (Manganese Oxide);120A36;30B;2D;1920;1920;2",
"Bread Crumbs;23B35PQ;30A;1A,1B,1C;320;400;0.6",
"Brewer's Grain, spent, dry;22C45;30A;1A,1B,1C;220;480;0.5",
"Brewer’s Grain, spent, wet;58C45T;30A;2A,2B;880;960;0.8",
"Brick, Ground, 1⁄8 inch;110B37;15;3D;1600;1920;2.2",
"Bronze Chips;40B45;30A;2D;480;800;2",
"Buckwheat;40B25N;45;1A,1B,1C;590;670;0.4",
"Calcine, Flour;80A35;30A;1A,1B,1C;1200;1360;0.7",
"Calcium Carbide;80D25N;30A;2D;1120;1440;2",
"Calcium Hydrate (Lime, Hydrated);40B35LM;30A;2D;640;640;0.8",
"Calcium Hydroxide (Lime, Hydrated);40B35LM;30A;2D;640;640;0.8",
"Calcium Lactate;28D45QTR;30A;2A,2B;420;460;0.6",
"Calcium Oxide (Lime, unslaked);63B35U;30A;1A,1B,1C;960;1040;0.6",
"Calcium Phosphate;45A45;30A;1A,1B,1C;640;800;1.6",
"Canola Meal (Rape Seed Meal)**;38;?;?;540;660;0.8",
"Carborundum;100D27;15;3D;1600;1600;3",
"Casein;36B35;30A;2D;580;580;1.6",
"Cashew Nuts;35C45;30A;2D;510;590;0.7",
"Cast Iron, Chips;165C45;30A;2D;2080;3200;4",
"Caustic Soda (Sodium Hydroxide);88B35RSU;30A;3D;1410;1410;1.8",
"Caustic Soda, Flakes;47C45RSUX;30A;3A,3B;750;750;1.5",
"Celite (Diatomaceous Earth);14A36Y;30B;3D;180;270;1.6",
"Cellulose + TBA;VTK;30B;2D;960;800;1.6",
"Cement, Aerated (Portland);68A16M;30B;2D;960;1200;1.4",
"Cement, Clinker;85D36;30B;3D;1200;1520;1.8",
"Cement, Mortar;133B35Q;30A;3D;2130;2130;3",
"Cement, Portland;94A26M;30B;2D;1510;1510;1.4",
"Cerrusite (Lead Carbonate);250A35R;30A;2D;3840;4160;1",
"Chalk, Crushed;85D25;30A;2D;1200;1520;1.9",
"Chalk, Pulverized;71A25MXY;45;2D;1070;1200;1.4",
"Charcoal, Ground;23A45;30A;2D;290;450;1.2",
"Charcoal, Lumps;23D45Q;30A;2D;290;450;1.4",
"Chocolate, Cake Pressed;43D25;30A;2B;640;720;1.5",
"Chrome Ore;133D36;30B;3D;2000;2240;2.5",
"Cinders, Blast Furnace;57D36T;30B;3D;910;910;1.9",
"Cinders, Coal;40D36T;30B;3D;640;640;1.8",
"Clay (Marl);80D36;30B;2D;1280;1280;1.6",
"Clay, Brick, Dry, Fines;110C36;30B;3D;1600;1920;2",
"Clay, Calcined;90B36;30B;3D;1280;1600;2.4",
"Clay, Ceramic, Dry, Fines;70A35P;30A;1A,1B,1C;960;1280;1.5",
"Clay, Dry, Lumpy;68D35;30A;2D;960;1200;1.8",
"Clinker, Cement (Cement Clinker);85D36;30B;3D;1200;1520;1.8",
"Clover Seed;47B25N;45;1A,1B,1C;720;770;0.4",
"Coal, Anthracite (River & Culm);58B35TY;30A;2A,2B;880;980;1",
"Coal, Anthracite, Sized, 1⁄2 inch;55C25;45;2A,2B;780;980;1",
"Coal, Bituminous, Mined;50D35LNYX;30A;1A,1B;640;960;1",
"Coal, Bituminous, Mined, Sized;48D35QV;30A;1A,1B;720;800;1",
"Coal, Bituminous, Mined, Slack;47C45T;30A;2A,2B;690;800;0.9",
"Coal, Lignite;41D35T;30A;2D;590;720;1",
"Cocoa Beans;38C25Q;30A;1A,1B;480;720;0.5",
"Cocoa, Nibs;35C25;45;2D;560;560;0.5",
"Cocoa, Powdered;33A45XY;30A;1B;480;560;0.9",
"Coconut, Shredded;2.1E+46;30B;2B;320;350;1.5",
"Coffee, Chaff;20B25FZMY;45;1A,1B;320;320;1",
"Coffee, Green Bean;29C25PQ;45;1A,1B;400;510;0.5",
"Coffee, Ground, Dry;25A35P;30A;1A,1B;400;400;0.6",
"Coffee, Ground, Wet;40A45X;30A;1A,1B;560;720;0.6",
"Coffee, Roasted Bean;25C25PQ;45;1B;320;480;0.4",
"Coffee, Soluble;19A35PUY;30A;1B;300;300;0.4",
"Coke, Breeze;30C37;15;3D;400;560;1.2",
"Coke, Loose;30D37;15;3D;400;560;1.2",
"Coke, Petrol, Calcined;40D37;15;3D;560;720;1.3",
"Compost;40D45TV;30A;3A,3B;480;800;1",
"Concrete, Pre-Mix,;103C36U;30B;3D;1360;1920;3",
"Copper Ore;135D36;30B;3D;1920;2400;4",
"Copper Ore, Crushed;125D36;30B;3D;1600;2400;4",
"Copper Sulphate, (Bluestone, Cupric Sulphate);85C35S;30A;2A,2B,2C;1200;1520;1",
"Copperas (Ferrous Sulphate);63C35U;30A;2D;800;1200;1",
"Copra, Cake Ground;43B45HW;30A;1A,1B,1C;640;720;0.7",
"Copra, Cake, Lumpy;28D35HW;30A;2A,2B,2C;400;480;0.8",
"Copra, Lumpy;22E35HW;30A;2A,2B,2C;350;350;1",
"Copra, Meal;43B35HW;30A;2D;640;720;0.7",
"Cork, Fine Ground;10B35JNY;30A;1A,1B,1C;80;240;0.5",
"Cork, Granulated;14C35JY;30A;1A,1B,1C;190;240;0.5",
"Corn Cobs, Ground;17C25Y;45;1A,1B,1C;270;270;0.6",
"Corn Fiber, Dry;14B46P;30B;1A,1B,1C;190;240;1",
"Corn Fiber, Wet;33B46P;30B;1A,1B,1C;240;800;1.5",
"Corn Oil, Cake;25D45HW;30A;1A,1B;400;400;0.6",
"Corn, Cracked;45B25P;45;1A,1B,1C;640;800;0.7",
"Corn, Germ, Dry;21B35PY;30A;1A,1B,1C;340;340;0.4",
"Corn, Germ, Wet (50%, moisture);30B35PY;30A;1A,1B,1C;480;480;0.4",
"Corn, Grits;43B35P;30A;1A,1B,1C;640;720;0.5",
"Corn, Seed;45C25PQ;45;1A,1B,1C;720;720;0.4",
"Corn, Shelled;45C25;45;1A,1B,1C;720;720;0.4",
"Corn, Starch*;38A15MN;45;1A,1B,1C;400;800;1",
"Corn, Sugar;33B35PU;30A;1B;480;560;1.3",
"Cornmeal;36B35P;30A;1A,1B;510;640;0.5",
"Cottonseed, Cake;43C45HW;30A;1A,1B;640;720;1",
"Cottonseed, Dry, Delinted;31C25X;45;1A,1B;350;640;0.6",
"Cottonseed, Dry, Not Delinted;22C45XY;30A;1A,1B;290;400;0.9",
"Cottonseed, Flakes;23C35HWY;30A;1A,1B;320;400;0.8",
"Cottonseed, Hulls;12B35Y;30A;1A,1B;190;190;0.9",
"Cottonseed, Meal, Expeller;28B45HW;30A;3A,3B;400;480;0.5",
"Cottonseed, Meal, Extracted;38B45HW;30A;1A,1B;560;640;0.5",
"Cottonseed, Meats, Dry;40B35HW;30A;1A,1B;640;640;0.6",
"Cottonseed, Meats, Rolled;38C45HW;30A;1A,1B;560;640;0.6",
"Cracklings, Crushed;45D45HW;30A;2A,2B,2C;640;800;1.3",
"Cryolite, Dust (Sodium Aluminum Fluoride);83A36V;30B;2D;1200;1440;2",
"Cryolite, Lumpy (Kryalith);100D36;30B;2D;1440;1760;2.1",
"Cullet, Fine;100C37;15;3D;1280;1920;2",
"Cullet, Lump;100D37;15;3D;1280;1920;2.5",
"Culm, (Coal, Anthracite);58B35TY;30A;2A,2B;880;980;1",
"Cupric Sulphate (Copper Sulfate);85C35S;30A;2A,2B,2C;1200;1520;1",
"Diatomaceous Earth (Filter Aid, Precoat);14A36Y;30B;3D;180;270;1.6",
"Dicalcium Phosphate;45A35;30A;1A,1B,1C;640;800;1.6",
"Disodium Phosphate;28A35;30A;3D;400;500;0.5",
"Distiller’s Grain, Spent Wet;50C45V;30A;3A,3B;640;960;0.8",
"Distiller’s Grain, Spent Wet w/Syrup;56C45VXOH;30A;3A,3B;690;1090;1.2",
"Distiller’s Grain-Spent Dry;30B35;30A;2D;480;480;0.5",
"Dolomite, Crushed;90C36;30B;2D;1280;1600;2",
"Dolomite, Lumpy;95D36;30B;2D;1440;1600;2",
"Earth, Loam, Dry, Loose;76C36;30B;2D;1220;1220;1.2",
"Ebonite, Crushed;67C35;30A;1A,1B,1C;1010;1120;0.8",
"Egg Powder;16A35MPY;30A;1B;260;260;1",
"Epsom Salts (Magnesium Sulfate);45A35U;30A;1A,1B,1C;640;800;0.8",
"Feldspar, Ground;73A37;15;2D;1040;1280;2",
"Feldspar, Lumps;95D37;15;2D;1440;1600;2",
"Feldspar, Powder;100A36;30B;2D;1600;1600;2",
"Felspar, Screenings;78C37;15;2D;1200;1280;2",
"Ferrous Sulfide, 1⁄2 inch (Iron Sulfide, Pyrites);128C26;30B;1A,1B,1C;1920;2160;2",
"Ferrous Sulfide, 100M (Iron Sulfide, Pyrites);113A36;30B;1A,1B,1C;1680;1920;2",
"Ferrous Sulphate (Iron Sulphate, Copperas);63C35U;30A;2D;800;1200;1",
"Filter-Aid (Diatomaceous Earth, Precoat);14A36Y;30B;3D;180;270;1.6",
"Fish Meal;38C45HP;30A;1A,1B,1C;560;640;1",
"Fish Scrap;45D45H;30A;2A,2B,2C;640;800;1.5",
"Flaxseed;44B35X;30A;1A,1B,1C;690;720;0.4",
"Flaxseed Cake (Linseed Cake);49D45W;30A;2A,2B;770;800;0.7",
"Flaxseed Meal (Linseed Meal);35B45W;30A;1A,1B;400;720;0.4",
"Flour Wheat;37A45LP;30A;1B;530;640;0.6",
"Flue Dust, Basic Oxygen Furnace;53A36LM;30B;3D;720;960;3.5",
"Flue Dust, Blast Furnace;118A36;30B;3D;1760;2000;3.5",
"Flue Dust, Boiler H. Dry;38A36LM;30B;3D;480;720;2",
"Fluorspar, Fine (Calcium Floride);90B36;30B;2D;1280;1600;2",
"Fluorspar, Lumps;100D36;30B;2D;1440;1760;2",
"Flyash;38A36M;30B;3D;480;720;2",
"Foundry Sand, Dry (Sand);95D37Z;15;3D;1440;1600;2.6",
"Fuller’s Earth, Calcined;40A25;45;3D;640;640;2",
"Fuller’s Earth, Dry, Raw (Bleach Clay);35A25;45;2D;480;640;2",
"Fuller’s Earth, Oily, Spent (Spent Bleach Clay);63C45OW;30A;3D;960;1040;2",
"Galena (Lead Sulfide);250A35R;30A;2D;3840;4160;5",
"Gelatine, Granulated;32B35PU;30A;1B;510;510;0.8",
"Gilsonite;37C35;30A;3D;590;590;1.5",
"Glass, Batch;90C37;15;3D;1280;1600;2.5",
"Glue, Ground;40B45U;30A;2D;640;640;1.7",
"Glue, Pearl;40C35U;30A;1A,1B,1C;640;640;0.5",
"Glue, Veg. Powdered;40A45U;30A;1A,1B,1C;640;640;0.6",
"Gluten, Meal (Dry Corn Gluten);40B35P;30A;1B;640;640;0.6",
"Gluten, Meal (Dry Corn Gluten) Syral;40B35P;30A;1B;500;500;0.6",
"Gluten, Meal (Wet Corn Gluten);43B35OPH;30A;1B;690;690;2.2",
"Granite, Fine;85C27;15;3D;1280;1440;2.5",
"Grape, Pomace;18D45U;30A;2D;240;320;1.4",
"Graphite Flake (Plumago);40B25LP;45;1A,1B,1C;640;640;0.5",
"Graphite Flour;28A35LMP;30A;1A,1B,1C;450;450;0.5",
"Graphite Ore;70D35L;30A;2D;1040;1200;1",
"Guano Dry**;70C35;30A;3A,3B;1120;1120;2",
"Gypsum, Calcined (Plaster of Paris);58B35U;30A;2D;880;960;1.6",
"Gypsum, Calcined, Powdered (Plaster of Paris);70A35U;30A;2D;960;1280;2",
"Gypsum, Raw 1 inch(Calc.Sulfate, Plast.of Paris);75D25;30A;2D;1120;1280;2",
"Hay, Chopped **;10C35JY;30A;2A, 2B;130;190;1.6",
"Hexanedioic Acid (Adipic Acid);45A35;30A;2B;720;720;0.5",
"Hominy, Dry;43C25D;30A;1A,1B,1C;560;800;0.4",
"Hops, Spend, Dry;35D35;30A;2A,2B,2C;560;560;1",
"Hops, Spent, Wet;53D45V;30A;2A,2B;800;880;1.5",
"Ice, Crushed;40D35O;30A;2A,2B;560;720;0.4",
"Ice, Cubes;34D35O;30A;1B;530;560;0.4",
"Ice, Flaked**;43C35O;30A;1B;640;720;0.6",
"Ice, Shell;34D45O;30A;1B;530;560;0.4",
"Ilmenite Ore (Titanium Dioxide);150D37;15;3D;2240;2560;2",
"Iron Ore Concentrate;150A37;15;3D;1920;2880;2.2",
"Iron Oxide Pigment;25A36LMP;30B;1A,1B,1C;400;400;1",
"Iron Oxide, Millscale;75C36;30B;2D;1200;1200;1.6",
"Iron Sulphate (Ferrous Sulfate);63C35U;30A;2D;800;1200;1",
"Iron Vitriol (Ferrous Sulfate);63C35U;30A;2D;800;1200;1",
"Kafir (Corn);43C25;45;3D;640;720;0.5",
"Kaolin Clay;63D25;30A;2D;1010;1010;2",
"Kaolin Clay (Tale);49A35LMP;30A;2D;670;900;2",
"Lactose;32A35PU;30A;1B;510;510;0.6",
"Lead Arsenate;72A35R;30A;1A,1B,1C;1150;1150;1.4",
"Lead Carbonate (Cerrusite);250A35R;30A;2D;3840;4160;1",
"Lead Ore, 1⁄2 inch;205C36;30B;3D;2880;3680;1.4",
"Lead Ore, 1⁄8 inch;235B35;30A;3D;3200;4320;1.4",
"Lead Oxide (Red Lead, Litharge) 100 Mesh;90A35P;30A;2D;480;2400;1.2",
"Lead Oxide (Red Lead, Litharge) 200 Mesh;105A35LP;30A;2D;480;2880;1.2",
"Lignite (Coal Lignite);41D35T;30A;2D;590;720;1",
"Limanite, Ore, Brown;120C47;15;3D;1920;1920;1.7",
"Lime Hydrated (Calcium Hydrate, Hydroxide);40B35LM;30A;2D;640;640;0.8",
"Lime Pebble;55C25HU;45;2A,2B;850;900;2",
"Lime, Ground, Unslaked (Quicklime);63B35U;30A;1A,1B,1C;960;1040;0.6",
"Lime, Hydrated, Pulverized;36A35LM;30A;1A,1B;510;640;0.6",
"Limestone, Agricultural (Calcium Carbonate);68B35;30A;2D;1090;1090;2",
"Limestone, Crushed (Calcium Carbonate); 88D36;30B;2D;1360;1440;2",
"Limestone, Dust (Calcium Carbonate);75A46MY;30B;2D;880;1520;1.8",
"Lindane (Benzene Hexachloride);56A45R;30A;1A,1B,1C;900;900;0.6",
"Linseed (Flaxseed);44B35X;30A;1A,1B,1C;690;720;0.4",
"Lithopone;48A35MR;30A;1A,1B;720;800;1",
"Magnesium Chloride (Magnesite);33C45;30A;1A,1B,1C;530;530;1",
"Maize (Milo);43B15N;45;1A,1B,1C;640;720;0.4",
"Malt, Dry Whole;25C35N;30A;1A,1B,1C;320;480;0.5",
"Malt, Dry, Ground;25C35N;30A;1A,1B,1C;320;480;0.5",
"Malt, Meal;38B25P;30A;1A,1B,1C;580;640;0.4",
"Malt, Sprouts;14C35P;30A;1A,1B,1C;210;240;0.4",
"Manganese Dioxide**;78A35NRT;30A;2A,2B;1120;1360;1.5",
"Manganese Ore;133D37;15;3D;2000;2240;2",
"Manganese Oxide;120A36;30B;2D;1920;1920;2",
"Manganese Sulfate;70C37;15;3D;1120;1120;2.4",
"Marble, Crushed;88B37;15;3D;1280;1520;2",
"Marl (Clay);80D36;30B;2D;1280;1280;1.6",
"Meat, Ground;53E45HQTX;30A;2A;800;880;1.5",
"Meat, Scrap (W/bone);40E46H;30B;2B;640;640;1.5",
"Mica, Flakes;20B16MY;30B;2D;270;350;1",
"Mica, Ground;14B36;30B;2D;210;240;0.9",
"Mica, Pulverized;14A36M;30B;2D;210;240;1",
"Milk, Dried, Flake;6B35PUY;30A;1B;80;100;0.4",
"Milk, Malted;29A45PX;30A;1B;430;480;0.9",
"Milk, Powdered;33B25PM;45;1B;320;720;0.5",
"Milk, Sugar;32A35PX;30A;1B;510;510;0.6",
"Milk, Whole, Powdered;28B35PUX;30A;1B;320;580;0.5",
"Mill Scale (Steel);123E46T;30B;3D;1920;2000;3",
"Milo Maize (Kafir);43B15N;45;1A,1B,1C;640;720;0.4",
"Milo, Ground (Sorghum Seed, Kafir);34B25;45;1A,1B,1C;510;580;0.5",
"Molybdenite Powder;107B26;30B;2D;1710;1710;1.5",
"Motar, Wet**;150E46T;30B;3D;2400;2400;3",
"Mustard Seed;45B15N;45;1A,1B,1C;720;720;0.4",
"Naphthalene Flakes;45B35;30A;1A,1B,1C;720;720;0.7",
"Niacin (Nicotinic Acid);35A35P;30A;2D;560;560;0.8",
"Oat Hulls;10B35NY;30A;1A,1B,1C;130;190;0.5",
"Oats;26C25MN;45;1A,1B,1C;420;420;0.4",
"Oats, Crimped;23C35;30A;1A,1B,1C;300;420;0.5",
"Oats, Crushed;22B45NY;30A;1A,1B,1C;350;350;0.6",
"Oats, Flour;35A35;30A;1A,1B,1C;560;560;0.5",
"Oats, Rolled;22C35NY;30A;1A,1B,1C;300;380;0.6",
"Oleo (Margarine);59E45HKPWX;30A;2A,2B;950;950;0.4",
"Orange Peel, Dry;1.5E+46;30A;2A,2B;240;240;1.5",
"Oxalic Acid Crystals – Ethane Diacid Crystals;60B35QS;30A;1A,1B;960;960;1",
"Oyster Shells, Ground;55C36T;30B;3D;800;960;1.8",
"Oyster Shells, Whole;80D36TV;30B;3D;1280;1280;2.3",
"Paper Pulp (4% or less);6.2E+46;30A;2A,2B;990;990;1.5",
"Paper Pulp (6% to 15%);6.2E+46;30A;2A,2B;960;990;1.5",
"Paraffin Cake, 1⁄2 inch;45C45K;30A;1A,1B;720;720;0.6",
"Peanut Meal;30B35P;30A;1B;480;480;0.6",
"Peanuts, Clean, in shell;18D35Q;30A;2A,2B;240;320;0.6",
"Peanuts, Raw (Uncleaned, Unshelled);18D36Q;30B;3D;240;320;0.7",
"Peanuts, Shelled;40C35Q;30A;1B;560;720;0.4",
"Peas, Dried;48C15NQ;45;1A,1B,1C;720;800;0.5",
"Perlite, Expanded;10C36;30B;2D;130;190;0.6",
"Phosphate Acid Fertilizer;60B25T;45;2A,2B;960;960;1.4",
"Phosphate Disodium (Sodium Phosphate);55A35;30A;1A,1B;800;960;0.9",
"Phosphate Rock, Broken;80D36;30B;2D;1200;1360;2.1",
"Phosphate Rock, Pulverized;60B36;30B;2D;960;960;1.7",
"Phosphate Sand;95B37;15;3D;1440;1600;2",
"Polyethylene, Resin Pellets;33C45Q;30A;1A,1B;480;560;0.4",
"Polystyrene Beads;40B35PQ;30A;1B;640;640;0.4",
"Polyvinyl Chloride Powder (PVC);25A45KT;30A;2B;320;480;1",
"Polyvinyl, Chloride Pellets;25E45KPQT;30A;1B;320;480;0.6",
"Potash (Muriate) Dry;70B37;15;3D;1120;1120;2",
"Potash (Muriate) Mine Run;75D37;15;3D;1200;1200;2.2",
"Potassium Carbonate;51B36;30B;2D;820;820;1",
"Potassium Nitrate, 1⁄2 inch (Saltpeter);76C16NT;30B;3D;1220;1220;1.2",
"Potassium Nitrate, 1⁄8 inch (Saltpeter);80B26NT;30B;3D;1280;1280;1.2",
"Potassium Sulfate;45B46X;30B;2D;670;770;1",
"Potassium-Chloride Pellets;125C25TU;45;3D;1920;2080;1.6",
"Potato Flour;48A35MNP;30A;1A,1B;770;770;0.5",
"PTA Crystal Slurry;VTK;--;--,--;--;1100;2.0",             'Toegevoegd 18-02-2016
"Pumice, 1⁄8 inch;45B46;30B;3D;670;770;1.6",
"Pyrite, Pellets;125C26;30B;3D;1920;2080;2",
"Quartz, 1⁄2 inch (Silicon Dioxide);85C27;15;3D;1280;1440;2",
"Quartz,100 Mesh (Silicon Dioxide);75A27;15;3D;1120;1280;1.7",
"Rape Seed Meal (Canola);38;?;?;540;660;0.8",
"Rice, Bran;20B35NY;30A;1A,1B,1C;320;320;0.4",
"Rice, Grits;44B35P;30A;1A,1B,1C;670;720;0.4",
"Rice, Hulled;47C25P;45;1A,1B,1C;720;780;0.4",
"Rice, Hulls;21B35NY;30A;1A,1B,1C;320;340;0.4",
"Rice, Polished;30C15P;45;1A,1B,1C;480;480;0.4",
"Rice, Rough;34C35N;30A;1A,1B,1C;510;580;0.6",
"Rosin, 1⁄2 inch;67C45Q;30A;1A,1B,1C;1040;1090;1.5",
"Rubber, Pelleted;53D45;30A;2A,2B,2C;800;880;1.5",
"Rubber, Reclaimed Ground;37C45;30A;1A,1B,1C;370;800;0.8",
"Rye;45B15N;45;1A,1B,1C;670;770;0.4",
"Rye Bran;18B35Y;30A;1A,1B,1C;240;320;0.4",
"Rye Feed;33B35N;30A;1A,1B,1C;530;530;0.5",
"Rye Meal;38B35;30A;1A,1B,1C;560;640;0.5",
"Rye Middlings;42B35;30A;1A,1B;670;670;0.5",
"Rye, Shorts;33C35;30A;2A,2B;510;530;0.5",
"Safflower Seed (Saffron);45B15N;45;1A,1B,1C;720;720;0.4",
"Safflower, Cake (Saffron);50D26;30B;2D;800;800;0.6",
"Safflower, Meal (Saffron);50B35;30A;1A,1B,1C;800;800;0.6",
"Sal Ammoniac (Ammonium Chloride);49A45FRS;30A;1A,1B,1C;720;830;0.7",
"Salicylic Acid;29B37U;15;3D;460;460;0.6",
"Salt Cake, Dry Coarse (Sodium Sulfate);85B36TU;30B;3D;1360;1360;2.1",
"Salt Cake, Dry Pulverized (Sodium Sulfate);75B36TU;30B;3D;1040;1360;1.7",
"Salt, Dry Coarse (Sodium Chloride);53C36TU;30B;3D;720;960;1",
"Salt, Dry Fine (Sodium Chloride);75B36TU;30B;3D;1120;1280;1.7",
"Sand (Resin Coated) Silica;104B27;15;3D;1670;1670;2",
"Sand (Resin Coated) Zircon;115A27;15;3D;1840;1840;2.3",
"Sand Dry Bank (Damp);120B47;15;3D;1760;2080;2.8",
"Sand Dry Bank (Dry);100B37;15;3D;1440;1760;1.7",
"Sand Dry Silica;95B27;15;3D;1440;1600;2",
"Sand Foundry (Shake Out);95D37Z;15;3D;1440;1600;2.6",
"Sawdust, Dry;12B45UX;30A;1A,1B,1C;160;210;0.7",
"Sea-Coal;65B36;30B;2D;1040;1040;1",
"Sesame Seed;34B26;30B;2D;430;660;0.6",
"Shale, Crushed;88C36;30B;2D;1360;1440;2",
"Shellac, Powdered or Granulated;31B35P;30A;1B;500;500;0.6",
"Silica Gel, 1⁄2 to 3 inch;45D37HKQU;15;3D;720;720;2",
"Silica, Flour;80A46;30B;2D;1280;1280;1.5",
"Slag, Blast Furnace Crushed;155D37Y;15;3D;2080;2880;2.4",
"Slag, Furnace Granular, Dry;63C37;15;3D;960;1040;2.2",
"Slate, Crushed, 1⁄2 inch;85C36;30B;2D;1280;1440;2",
"Slate, Ground, 1⁄8 inch;84B36;30B;2D;1310;1360;1.6",
"Sludge, Sewage, Dried;45E47TW;15;3D;640;800;0.8",
"Sludge, Sewage, Dry Ground;50B46S;30B;2D;720;880;0.8",
"Soap Detergent;33B35FQ;30A;1A,1B,1C;240;800;0.8",
"Soap, Beads or Granules;25B35Q;30A;1A,1B,1C;240;560;0.6",
"Soap, Chips;20C35Q;30A;1A,1B,1C;240;400;0.6",
"Soap, Flakes;10B35QXY;30A;1A,1B,1C;80;240;0.6",
"Soap, Powder;23B25X;45;1A,1B,1C;320;400;0.9",
"Soapstone, Talc, Fine;45A45XY;30A;1A,1B,1C;640;800;2",
"Soda Ash, Heavy (Sodium Carbonate);60B36;30B;2D;880;1040;1",
"Soda Ash, Light (Sodium Carbonate);28A36Y;30B;2D;320;560;0.8",
"Sodium Aluminate, Ground;72B36;30B;2D;1150;1150;1",
"Sodium Aluminum Sulphate**;75A36;30B;2D;1200;1200;1",
"Sodium Bicarbonate (Baking Soda);48A25;45;1B;640;880;0.6",
"Sodium Nitrate;75D25NS;30A;2A,2B;1120;1280;1.2",
"Sodium Phosphate;55A35;30A;1A,1B;800;960;0.9",
"Sodium Sulfite;96B46X;30B;2D;1540;1540;1.5",
"Soybean Meal Hot;40B35T;30A;2A,2B;640;640;0.5",
"Soybean Meal, Cold;40B35;30A;1A,1B,1C;640;640;0.5",
"Soybean, Cake;42D35W;30A;2A,1B,1C;640;690;1",
"Soybean, Cracked;35C36NW;30B;2D;480;640;0.5",
"Soybean, Flake, Extracted, Wet;34C35;30A;1A,1B,1C;540;540;0.8",
"Soybean, Flake, Raw;22C35Y;30A;1A,1B,1C;290;400;0.8",
"Soybean, Flour;29A35MN;30A;1A,1B,1C;430;480;0.8",
"Soybeans, Whole;48C26NW;30B;3D;720;800;1",
"Starch;38A15M;45;1A,1B,1C;400;800;1",
"Steel Turnings, Crushed;125D46WV;30B;3D;1600;2400;3",
"Sugar Beet, Pulp, Dry;14C26;30B;2D;190;240;0.9",
"Suga Beet, Pulp, Wet;35C35X;30A;1A,1B,1C;400;720;1.2",
"Sugar, Powdered;55A35PX;30A;1B;800;960;0.8",
"Sugar, Raw;60B35PX;30A;1B;880;1040;1.5",
"Sugar, Refined, Granulated Dry;53B35PU;30A;1B;800;880;1.2",
"Sugar, Refined, Granulated Wet;60C35P;30A;1B;880;1040;2",
"Sulphur, Crushed, 1⁄2 inch;55C35N;30A;1A,1B;800;960;0.8",
"Sulphur, Lumpy, 3 inch;83D35N;30A;2A,2B;1280;1360;0.8",
"Sulphur, Powdered;55A35MN;30A;1A,1B;800;960;0.6",
"Sunflower Seed;29C15;45;1A,1B,1C;300;610;0.5",
"Sunflower Seed Flakes;28C35;30A;1A,1B,1C;430;450;0.8",
"Swee Bran Feed (proprietary to Cargill);29B45P;30A;1A,1B,1C;340;590;0.6",
"Talcum Powder;55A36M;30B;2D;800;960;0.8",
"Talcum, 1⁄2 ich;85C36;30B;2D;1280;1440;0.9",
"Tanbark, Ground**;55B45;30A;1A,1B,1C;880;880;0.7",
"Timothy Seed;36B35NY;30A;1A,1B,1C;580;580;0.6",
"Titanium Dioxide based pigments (powder);42C36FLO;15;3D;540;800;2",
"Tobacco, Scraps;20D45Y;30A;2A,2B;240;400;0.8",
"Tobacco, Snuff;30B45MQ;30A;1A,1B,1C;480;480;0.9",
"Tricalcium Phosphate;45A45;30A;1A,1B;640;800;1.6",
"Triple Sugar Phosphate ;53B36RS;30B;3D;800;880;2",
"Trisodium Phosphate;60C36;30B;2D;960;960;1.7",
"Trisodium Phosphate Granular;60B36;30B;2D;960;960;1.7",
"Trisodium Phosphate, Pulverized;50A36;30B;2D;800;800;1.6",
"Tung Nut Meats, Crushed;28D25W;30A;2A,2B;450;450;0.8",
"Tung Nuts ;28D15;30A;2A,2B;400;480;0.7",
"Urea Prills, Coated;45B25;45;1A,1B,1C;690;740;1.2",
"Vermiculite, Expanded;16C35Y;30A;1A,1B;260;260;0.5",
"Vermiculite, Ore;80D36;30B;2D;1280;1280;1",
"Vetch;48B16N;30B;1A,1B,1C;770;770;0.4",
"Walnut Shells, Crushed;40B36;30B;2D;560;720;1",
"Wheat;47C25N;45;1A,1B,1C;720;770;0.4",
"Wheat Flour;37A45LP;45;1B;530;640;0.6",
"Wheat, Cracked;43B25N;45;1A,1B,1C;640;720;0.4",
"Wheat, Germ;23B25;45;1A,1B,1C;290;450;0.4",
"White Lead, Dry;88A36MR;30B;2D;1200;1600;1",
"Wood Chips, Screened;20D45VY;30A;2A,2B;160;480;0.6",
"Wood Flour;26B35N;30A;1A,1B;260;580;0.4",
"Wood Shavings;12E45VY;30A;2A,2B;130;260;1.5",
"Zinc Oxide, Heavy;33A45X;30A;1A,1B;480;560;1",
"Zinc Oxide, Light;13A45XY;30A;1A,1B;160;240;1",
"Zinc, Concentrate Residue;78B37;15;3D;1200;1280;1"}

    '--- "Oude benaming;Norm:;EN10027-1;Werkstof;[mm/m1/100°C];Poisson ;kg/m3;E [Gpa];Rm (20c);Rp0.2(0c);Rp0.2(20c);Rp(50c);Rp(100c);Rp(150c);Rp(200c);Rp(250c);Rp(300c);Rp(350c);Rp(400c);Equiv-ASTM;Opmerking",
    Public Shared steel() As String =
     {"Oude benaming;Norm:;EN10027-1;Werkstof;[mm/m1/100°C];Poisson ;kg/m3;E [Gpa];Rm (20c);0;20;50;100;150;Rp(200c);Rp(250c);Rp(300c);Rp(350c);Rp(400c);Equiv-ASTM;Opmerking",
      "Domex 690XPD(E);EN10149-2 UNS;S700MCD(E);1.8974;1.29;0.28;7850;192;810;675;740;765;690;675;660;640;620;580;540;--;--",
    "Duplex(Avesta-2205);EN 10088-1 UfllW;X2CrNiMoN22-5-3 saisna;1.4462;1.4;0.28;7800;200;640-950;335;460;385;360;335;315;300;0;0;0;A240-S31803;Max 300c",
    "Hastelloy-C22;DIN Nr: ASTM UNS;NiCr21Mo14W 2277 B575 N06022;2.4602;1.25;0.29;9000;205;786-800;310;370;354;338;310;283;260;248;244;241;--;--",
    "Inconel- 600;DIN Nicrofer7216 ASTM SO ;NiCr15Fe Alloy 600 B168 NiCr15Fe8 Npsepo;2.4816;1.44;0.29;8400;214;550;170;240;185;180;170;165;160;155;152;150;--;--",
    "P265GH;EN 10028-2 UNS;P265GH ;1.0425;1.29;0.28;7850;192;410-530;205;255;234;215;205;195;175;155;140;130;A516-Gr60;--",
    "S235JRG2;EN 10025 UNS;S235JRG2 ;1.0038;1.29;0.28;7850;192;340-470;180;195;200;190;180;170;150;130;120;110;A283-GrC;--",
    "S355J2G3;EN10025 UNS;S355J2G3;1.057;1.29;0.28;7850;192;490-630;284;315;340;304;284;255;226;206;0;0;A299;Max 300c	",
    "SS 304;EN10088-2;X5CrNI18-10 S30400;1.4301;1.76;0.28;7900;200;520-750;142;210;165;157;142;127;118;110;104;98;A240-304;--",
    "SS 304L;EN10088-2;X2CrNi19-11 S30403;1.4306;1.76;0.28;7900;200;520-670;132;200;155;147;132;118;108;100;94;89;A240-304L;--",
    "SS 316;EN10088-2;X5CrNiMo17-12-2 S31600;1.4401;1.76;0.28;8000;200;520-680;162;220;180;177;162;147;137;127;120;115;A240-316;--",
    "SS 316L;EN10088-2;X2CrNiMo17-12-2 S31603;1.4404;1.76;0.28;8000;200;520-680;152;220;170;166;152;137;127;118;113;108;A240-316L;--",
    "SS 316TI;EN10088-2;X6CrNiMoTi17-12-2 S31635;1.4571;1.76;0.28;8000;200;520-690;177;220;191;185;177;167;157;145;140;135;A240-316Ti;--",
    "SS 321;EN10088-2;X6CrNiTi18-10 S32100;1.4541;1.76;0.28;7900;200;500-720;167;200;184;176;167;157;147;136;130;125;A240-321;--",
    "SS 410 ;EN 10088-1 U1S;X12Cr13 (Gegloeid) 541000;1.4006;1.15;0.28;7700;216;450-650;230;250;240;235;230;225;225;220;210;195;A240-410;--",
    "SuperDuplex;--;X2CrNiMoN22-5-3 saisna;1.4501;1.4;0.28;7800;200;730-930;445;550;510;480;445;405;400;395;0;0;--;--"}

    'DN, inch,OD, wall1, wall2, wall3,...
    Public Shared pipe_ss() As String =
   {"DN100;4 inch; 114.3;  6.3;7.1;8;10;12.5;16",
    "DN125;5 inch; 139.7;  6.3;7.1;8;10;12.5;16",
    "DN150;6 inch; 168.3;  6.3;7.1;8;10;12.5;16",
    "DN200;8 inch; 219.1;  6.3;7.1;8;10;12.5;16",
    "DN250;10 inch; 273;   6.3;7.1;8;10;12.5;16",
    "DN300;12 inch; 323.9; 6.3;7.1;8;10;12.5;16",
    "DN350;14 inch; 355.6; 6.3;7.1;8;10;12.5;16",
    "DN400;16 inch; 406.4; 6.3;7.1;8;10;12.5;16",
    "DN500;20 inch; 508;   6.3;7.1;8;10;12.5;16"}

    Public Shared pipe() As String =
   {"DN100;4 inch; 114.3;  6.02;  8.56; -;   0",
    "DN125;5 inch; 141.3;  6.55;  9.53; -;   0",
    "DN150;6 inch; 168.3;  7.11; 10.97; -;   0",
    "DN200;8 inch; 219.1;  6.35;  8.18; 12.7;0",
    "DN250;10 inch; 273;   6.35;  9.27; 12.7;0",
    "DN300;12 inch; 323.9; 6.35;  9.27; 12.7;0",
    "DN350;14 inch; 355.6; 7.92;  9.53; -;   0",
    "DN400;16 inch; 406.4; 7.92;  9.53; -;   0"}

    Public Shared emotor() As String = {"3.0; 1500", "4.0; 1500", "5.5; 1500", "7.5; 1500", "11;  1500", "15; 1500", "22; 1500",
                                       "30  ; 1500", "37;  1500", "45;  1500", "55;  1500", "75; 1500", "90; 1500",
                                       "110 ; 1500", "132; 1500", "160; 1500", "200; 1500"}

    Public Shared diam_trough As Double                         '[m]
    Public Shared pipe_OD, pipe_ID, pipe_wall As Double
    Public Shared pipe_Ix, pipe_Wx, pipe_Wp As Double            'Lineair en polair weerstand moment
    Public Shared pitch As Double
    Public Shared installed_power As Double
    Public Shared sigma02, sigma_fatique, Elast As Double
    Public Shared inlet_length, conv_length, product_density As Double
    Dim angle As Double
    Dim speed As Double
    Dim flow_hr As Double
    Dim density As Double
    Dim filling_perc As Double
    Dim progress_resistance As Double                   'Friction from product to steel

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")


        For hh = 0 To (UBound(inputs) - 1)              'Fill combobox1
            words = inputs(hh).Split(";")
            ComboBox1.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 225                   'Grafite ore

        '-------Fill combobox2, Steel selection------------------
        For hh = 0 To (UBound(steel) - 1)               'Fill combobox 2 with steel data
            words = steel(hh).Split(";")
            ComboBox2.Items.Add(words(0))
        Next hh
        ComboBox2.SelectedIndex = 8                     'rvs 304


        '-------Fill combobox5, emotor selection------------------
        For hh = 0 To (UBound(emotor) - 1)               'Fill combobox 5 emotor data
            words = emotor(hh).Split(";")
            ComboBox5.Items.Add(words(0))
        Next hh
        ComboBox5.SelectedIndex = 0

        pipe_dia_combo()
        pipe_wall_combo()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, Button1.Click, NumericUpDown1.ValueChanged, TabPage1.Enter
        calculate()
        calulate_stress_1()
    End Sub

    Private Sub calculate()
        Dim cap_hr As Double    '100% Çapacity conveyor [m3/hr]
        Dim cap_act As Double   'actual Çapacity conveyor [m3/hr]
        Dim iso_forward As Double       'Power for forward motion
        Dim iso_incline As Double       'Power inclination
        Dim iso_no_product As Double    'Power for seals + bearings
        Dim iso_power As Double         'Total Power 
        Dim height As Double            'Height difference due to inclination 
        Dim mekog As Double             'Mekog installed power
        Dim flight_speed As Double         'Flight speed

        '-------------- get data----------
        diam_trough = NumericUpDown1.Value / 1000               'Trough width[m]
        TextBox18.Text = diam_trough * 1000.ToString
        TextBox16.Text = pipe_OD * 1000.ToString                'Pipe diameter [m]

        pitch = diam_trough * NumericUpDown2.Value              '[-]
        conv_length = NumericUpDown3.Value                      'Conveyor length [m]
        TextBox19.Text = conv_length.ToString

        angle = NumericUpDown4.Value                            '[degree]
        speed = NumericUpDown7.Value                            '[rpm]

        Flight_speed = speed / 60 * PI * diam_trough
        TextBox11.Text = Round(flight_speed, 2).ToString 'Flight speed [m/s]

        If flight_speed > 1.0 Then
            TextBox11.BackColor = Color.Red
        Else
            TextBox11.BackColor = Color.LightGreen
        End If

        If speed > 45 Then
            NumericUpDown7.BackColor = Color.Red
        Else
            NumericUpDown7.BackColor = Color.Yellow
        End If

        flow_hr = NumericUpDown5.Value * 1000           '[kg/hr]
        density = NumericUpDown6.Value                  '[kg/m3]
        progress_resistance = NumericUpDown9.Value      '[-]

        '--------------- now calc-----------------
        cap_hr = PI / 4 * (diam_trough ^ 2 - pipe_OD ^ 2) * pitch * speed * 60            ' [m]
        cap_hr = cap_hr * (100 - angle * 2) / 100                                         ' capacity loss due to inclination (2% per degree)

        cap_act = flow_hr / density
        filling_perc = Round(cap_act / cap_hr * 100, 1)

        If filling_perc > 40 Then
            TextBox1.BackColor = Color.Red
        Else
            TextBox1.BackColor = Color.LightGreen
        End If

        '--------------- ISO 7119 -----------------
        height = conv_length * Sin(angle / 360 * 2 * PI)

        iso_forward = flow_hr * conv_length * 9.91 * progress_resistance / (3600 * 1000)     'Forwards [kW]
        iso_incline = flow_hr * height * 9.81 / (3600 * 1000)                           'Uphill [kW]
        iso_no_product = diam_trough * conv_length / 20                                      'Power for seals 0. + bearings [kW]

        iso_power = Round(iso_forward + iso_incline + iso_no_product, 1)

        '--------------- MEKOG -----------------
        mekog = Round(flow_hr * conv_length / (40 * 1.36 * 1000), 1)    '[kW]

        '--------------- present results------------
        TextBox1.Text = filling_perc.ToString
        TextBox3.Text = iso_power.ToString
        TextBox4.Text = mekog.ToString
        'End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        pipe_wall_combo()   'Put new wall thicknesses in the combobox
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        save_to_disk()
    End Sub

    'Materiaal in de conveyor
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim words() As String = inputs(ComboBox1.SelectedIndex).Split(";")
            NumericUpDown6.Value = words(5) 'Density max
            NumericUpDown9.Value = words(6) 'Material factor
            Label37.Text = "CEMA material code " & words(1)
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    'Please note complete calculation in [m] nit [mm]
    Private Sub calulate_stress_1()
        Dim qq As Double
        Dim Q_load_1, Q_load_2, Q_load_comb, Q_load_3, Q_Deflect_max, Q_max_bend, pos_x As Double
        Dim F_tangent, Radius_transport As Double
        Dim weight_m As Double
        Dim pipe_OR, pipe_IR As Double
        Dim sigma_eg As Double                      'Sigma eigen gewicht
        Dim flight_hoogte, flight_gewicht, flight_lengte_buiten, flight_lengte_binnen, flight_lengte_gem, fligh_dik As Double
        Dim P_torque, Tou_torque As Double           'Torque @ aandrijving
        Dim P_torque_M, Tou_torque_M As Double       'Torque @ max bend
        Dim words() As String
        Dim Ra, Rb, R_total As Double
        Dim kolom_height As Double
        Dim combined_stress As Double
        Dim max_sag As Double                       'maximale doorzakking pijp

        NumericUpDown13.Value = NumericUpDown7.Value

        '---------- materiaal gewicht inlaat kolom op pipe--------------------------
        inlet_length = NumericUpDown16.Value                'inlet chute [m]
        kolom_height = NumericUpDown17.Value                'inlet chute material kolom [m]
        product_density = NumericUpDown6.Value              'product density [kg/m3]

        If (ComboBox5.SelectedIndex > -1) Then      'Prevent exceptions
            words = emotor(ComboBox5.SelectedIndex).Split(";")
            Double.TryParse(words(0), installed_power)
        End If

        If (ComboBox2.SelectedIndex > -1) Then      'Prevent exceptions
            words = steel(ComboBox2.SelectedIndex).Split(";")
            TextBox6.Text = words(6)     'Density steel

            '--------------- select the strength @ temperature
            qq = NumericUpDown11.Value
            Select Case True
                Case (qq > -10 AndAlso qq <= 0)
                    Double.TryParse(words(9), sigma02)     'Sigma 0.2 [N/mm]
                Case (qq > 0 AndAlso qq <= 20)
                    Double.TryParse(words(10), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 20 AndAlso qq <= 50)
                    Double.TryParse(words(11), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 50 AndAlso qq <= 100)
                    Double.TryParse(words(12), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 100 AndAlso qq <= 150)
                    Double.TryParse(words(13), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 150 AndAlso qq <= 200)
                    Double.TryParse(words(13), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 200 AndAlso qq <= 250)
                    Double.TryParse(words(14), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 250 AndAlso qq <= 300)
                    Double.TryParse(words(15), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 300 AndAlso qq <= 350)
                    Double.TryParse(words(16), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 350 AndAlso qq <= 400)
                    Double.TryParse(words(17), sigma02)    'Sigma 0.2 [N/mm]
            End Select
            TextBox7.Text = sigma02
            sigma_fatique = sigma02 * 0.35                   'Fatique stress uitgelegd op oneindige levensduur
            TextBox8.Text = Round(sigma_fatique, 0).ToString
        End If

        If ComboBox3.SelectedIndex > -1 Then
            words = pipe(ComboBox3.SelectedIndex).Split(";")

            Double.TryParse(words(2), pipe_OD)
            pipe_OD /= 1000                             'Outside Diameter [m]
            pipe_OR = pipe_OD / 2                       'Radius [mm]
            pipe_wall = ComboBox6.SelectedItem / 1000   'Wall thickness [mm]
            pipe_ID = (pipe_OD - 2 * pipe_wall)         'Inside diameter [mm]
            pipe_IR = pipe_ID / 2                       'Inside radius [mm]

            weight_m = PI / 4 * (pipe_OD ^ 2 - pipe_ID ^ 2) * 7850          'Weight per meter [kg/m]

            TextBox13.Text = Round(weight_m, 1).ToString                    'gewicht per meter
            TextBox16.Text = Round(pipe_OD * 1000, 1).ToString              'Diameter [m]

            '---------------- Traagheids moment Ix= PI/64.(D^4-d^4)---------------------
            pipe_Ix = PI / 64 * (pipe_OD ^ 4 - pipe_ID ^ 4)                  '[m4]
            TextBox26.Text = Round(pipe_Ix * 1000 ^ 4, 0).ToString

            '---------------- Weerstand moment Buiging  Wx= PI/32.(D^4-d^4)/D---------------------
            pipe_Wx = PI / 32 * (pipe_OD ^ 4 - pipe_ID ^ 4) / pipe_OD        '[m3]
            TextBox14.Text = Round(pipe_Wx * 1000 ^ 3, 0).ToString

            '---------------- Weerstand moment Torsie (polair)  Wp= PI/16.(D^4-d^4)/D --------------
            pipe_Wp = PI / 16 * (pipe_OD ^ 4 - pipe_ID ^ 4) / pipe_OD       '[m3]
            TextBox15.Text = Round(pipe_Wp * 1000 ^ 3, 0).ToString


            '============================================calc load ==============================================================================
            '====================================================================================================================================

            '---------------- gewicht flight---mm dik----------------------------------
            flight_hoogte = (diam_trough - pipe_OD / 1000) / 2                                  '[m]
            flight_lengte_buiten = Sqrt((PI * diam_trough) ^ 2 + (pitch) ^ 2)

            flight_lengte_binnen = Sqrt((PI * pipe_OD / 1000) ^ 2 + (pitch) ^ 2)
            flight_lengte_gem = (flight_lengte_buiten + flight_lengte_binnen) / 2
            fligh_dik = NumericUpDown8.Value / 1000                                             '[m]
            flight_gewicht = (flight_hoogte * flight_lengte_gem * fligh_dik * 7850) / pitch     'Flight Gewicht per meter
            TextBox2.Text = Round(flight_gewicht, 1).ToString                                   'Flight Gewicht per meter
            TextBox5.Text = Round(fligh_dik * 1000, 1).ToString                                 'Flight dikte [mm]

            '------------- aandrijving torsie @ drive ---------------------------------------------------------------------------------
            P_torque = installed_power * 1000 / (2 * PI * NumericUpDown7.Value / 60)
            TextBox29.Text = Round(P_torque, 0).ToString                                        'Torque from drive [N.m]


            '----------- Weight (pipe+flight) + transport force combined ------
            '---- Worst case material assumed sitting lowest point of the trough---

            Q_load_1 = (weight_m + flight_gewicht) * 9.81           'Total EVEN distributed load
            TextBox22.Text = Round(Q_load_1, 0).ToString            'Belasting [kg/m]

            '----------- Axial load caused by transport of product
            Radius_transport = (diam_trough + pipe_OD) / 4                  'Acc Jos (D+d)/4
            F_tangent = P_torque / Radius_transport
            Q_load_2 = F_tangent / conv_length                              'Transport kracht geeft doorbuiging pijp
            Q_load_3 = pipe_OD * kolom_height * product_density * 9.91      'gelijkmatige belasting op de pijp door materiaal kolom
            TextBox17.Text = Round(Q_load_3, 0).ToString                    '[N/m]

            '============================================ Traditionele VTK berekening ===========================================================
            '============================================ verwaarloosd Q_load2 =======================================================================
            If CheckBox1.Checked Then
                Q_load_2 = 0
            End If
            TextBox28.Text = Round(Q_load_2, 0).ToString                    '[N]
            '============================================ Reactie krachten ======================================================================
            '====================================================================================================================================

            Q_load_comb = Sqrt(Q_load_1 ^ 2 + Q_load_2 ^ 2)     'Radiale en tangentiele kracht gecombineerd


                R_total = Q_load_comb * conv_length + Q_load_3 * inlet_length    'Total force (Ra+Rb)

                'Momenten evenwicht om punt Ra
                Rb = (Q_load_comb * conv_length ^ 2 * 0.5 + Q_load_3 * inlet_length ^ 2 * 0.5) / conv_length
                Ra = R_total - Rb

                TextBox24.Text = Round(Ra, 0).ToString          'Reactie kracht Ra
                TextBox36.Text = Round(Rb, 0).ToString          'Reactie kracht Rb
                TextBox39.Text = Round(R_total, 0).ToString     'Reactie kracht Ra+Rb

                '============================================ Maximaal moment positie ===============================================================
                '=============================================  Maximaal moment (oppervlak dwarskrachtenlijn) =======================================
                pos_x = Ra / (Q_load_comb + Q_load_3)
                Q_max_bend = 0.5 ^ 2 * (Q_load_comb + Q_load_3) * pos_x

                If pos_x > inlet_length Then
                    pos_x = conv_length - Rb / Q_load_comb
                    Q_max_bend = 0.5 ^ 2 * Q_load_comb * (conv_length - pos_x)
                End If

                TextBox38.Text = Round(pos_x, 2).ToString           'Positie max moment tov A [m]
                TextBox37.Text = Round(Q_max_bend, 0).ToString      'Max moment [Nm]          

                '============================================calc torsie ============================================================================
                '====================================================================================================================================
                Tou_torque = P_torque / (pipe_Wp * 1000 ^ 2)            '[N/mm2]
                TextBox12.Text = Round(Tou_torque, 1).ToString          'Stress from drive [N.m]
                If Tou_torque > sigma_fatique Then
                    TextBox12.BackColor = Color.Red
                Else
                    TextBox12.BackColor = Color.LightGreen
                End If

                '-------------------------- @ max bend------------------------
                P_torque_M = (P_torque * pos_x / conv_length)
                Tou_torque_M = P_torque_M / (pipe_Wp * 1000 ^ 2)            '[N/mm2]
                TextBox10.Text = Round(Tou_torque_M, 1).ToString


                '============================================calc stress ============================================================================
                '====================================================================================================================================

                '----------- bending stress--------------------
                sigma_eg = Q_max_bend / (pipe_Wx * 1000 ^ 2)                   '[N/mm2]
                TextBox9.Text = Round(sigma_eg, 1).ToString                    '[N/mm2]

                If sigma_eg > sigma_fatique Then
                    TextBox9.BackColor = Color.Red
                Else
                    TextBox9.BackColor = Color.LightGreen
                End If

                '------------ Hubert en hencky @ maximale doorbuiging--------------------
                combined_stress = Sqrt((sigma_eg) ^ 2 + 3 * (Tou_torque_M) ^ 2)
                TextBox21.Text = Round(combined_stress, 1).ToString

                If combined_stress > sigma_fatique Then
                    TextBox21.BackColor = Color.Red
                Else
                    TextBox21.BackColor = Color.LightGreen
                End If


                '---------------- Max doorbuiging gelijkmatige belasting f= 5.Q.L^4/(384 .E.I) --------------------
                '---------------- materiaal kolom is niet meegenomen ----------------------------------------------
                Elast = 210 * 1000 ^ 2                                   '[N/mm2]

                Q_Deflect_max = (5 * Q_load_comb * conv_length ^ 4) / (384 * Elast * pipe_Ix)
                TextBox20.Text = Round(Q_Deflect_max, 1).ToString     '[mm]

                Select Case True
                    Case (RadioButton1.Checked)
                        max_sag = 500
                    Case (RadioButton2.Checked)
                        max_sag = 800
                    Case (RadioButton3.Checked)
                        max_sag = 1000
                End Select

                If Q_Deflect_max > conv_length * 1000 / max_sag Then
                    TextBox20.BackColor = Color.Red
                Else
                    TextBox20.BackColor = Color.LightGreen
                End If
            End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown11.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown13.ValueChanged, TabPage2.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, ComboBox3.ValueMemberChanged, ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, CheckBox1.CheckedChanged
        calulate_stress_1()
    End Sub
    Private Sub pipe_dia_combo()
        Dim words() As String

        ComboBox3.Items.Clear()

        '-------Fill combobox3, Pipe selection------------------
        For hh = 0 To (UBound(pipe) - 1)                'Fill combobox 3 with pipe data
            words = pipe(hh).Split(";")
            ComboBox3.Items.Add(Trim(words(2)))
        Next hh
        ComboBox3.SelectedIndex = 2

        words = pipe(ComboBox3.SelectedIndex).Split(";")
        Double.TryParse(words(2), pipe_OD)
        pipe_OD /= 1000                                         'Outside Diameter [m]
        TextBox16.Text = Round(pipe_OD * 1000, 1).ToString      'Diameter [mm]
    End Sub
    Private Sub pipe_wall_combo()
        Dim words() As String
        Dim temp As Double

        ComboBox6.Items.Clear()
        '-------Fill combobox6, pipe wall selection------------------
        words = pipe(ComboBox3.SelectedIndex).Split(";")  'Fill combobox 6 pipe wall data
        For hh = 3 To 5
            If Double.TryParse(words(hh), temp) Then
                ComboBox6.Items.Add(Trim(words(hh)))
            End If
        Next
        ComboBox6.SelectedIndex = 1
    End Sub

    'Save data and line chart to file
    Private Sub save_to_disk()
        Dim bmp_tab_page1 As New Bitmap(TabPage1.Width, TabPage1.Height)
        Dim bmp_tab_page2 As New Bitmap(TabPage2.Width, TabPage2.Height)
        Dim str_file2, str_file3 As String

        Dim text As String
        text = Now.ToString("yyyy_MM_dd_HH_mm_ss_")

        str_file2 = "c:\temp\" & text & "Conveyor selection data.png"
        str_file3 = "c:\temp\" & text & "Conveyor stress waaier.png"

        '---- save tab page 1---------------
        TabPage2.Show()
        TabPage1.DrawToBitmap(bmp_tab_page1, DisplayRectangle)
        bmp_tab_page1.Save(str_file2, Imaging.ImageFormat.Png)

        '---- save tab page 2---------------
        TabPage2.Show()
        TabPage2.DrawToBitmap(bmp_tab_page2, DisplayRectangle)
        bmp_tab_page2.Save(str_file3, Imaging.ImageFormat.Png)

        MessageBox.Show("Files is saved to c:\temp ")

    End Sub
End Class

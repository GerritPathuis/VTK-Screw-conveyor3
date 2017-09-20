Imports System.Math
Imports System.IO
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word


Public Class Form1
    'Materials name; CEMA Material code; Conveyor loading; Component group, density min, Density max, HP Material
    Public Shared _inputs() As String = {
"Adipic Acid;45A35;30A;2B;720;720;0.5",
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
"Cellulose with TBA;VTK;30B;2D;960;800;1.6",
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
"Hisarna Granulaat;--;--;--,--,--;--;2000;4",                       'Toegevoegd 25-6-2016
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


    Public Shared motorred() As String =
     {"Description; Speed; power;cost;shaftdia",
     "0.18 Kw,R27DR63M4;69.5;0.18;253.51;25",
     "3 Kw,Bauer BG60-11/DHE11XAC-TF;49.5;3;1132;50",
     "3 Kw, 20rpmR107;20;3;1908.74;70",
     "3 Kw, R77DRM100L4;29.12;3;896.25;40",
     "3 Kw, R137R77/II2GD EDRE100LC4;6.2;3;3851.01;90",
     "1.1 Kw, R77/II2GD EDRE90M4;27;1.1;814.50;40",
     "0.75 Kw, R87DRE90L6;940;0.75;1003.71;50",
     "2.2 Kw, R47DRE100M4;14.56;2.2;471.18;30",
      "1.1 Kw,R97DRN90S4;186;1.1;1340.06;60",
      "2.2 Kw 15.1023;55;2.2;0;0",
      "1.5 Kw, 1440 rpm;1440;1.5;0;0",
      "2.2 Kw, 1440 rpm;1440;2.2;0;0",
      "3 Kw, 1455 rpm;1455;3;0;0",
      "4 Kw, 1465 rpm;1465;4;0;0",
      "5.5 Kw, 1475 rpm;1475;5.5;0;0",
      "7.5 Kw, 1475 rpm;1475;7.5;0;0",
      "9.2 Kw, 1475 rpm;1475;9.2;0;0",
      "11 Kw, 1475 rpm;1475;11;0;0",
      "15 Kw, 1475 rpm;1475;15;0;0",
      "18.5 Kw, 1480 rpm;1480;18.5;0;0",
      "22 Kw, 1482 rpm;1482;22;0;0",
      "30 Kw, R137DRP225S4/TF/PT;53;30;5802.48;90",
      "45 Kw, R147DRP250M4/TF/NIB/PT;49;45;8588.74;110",
      "37 Kw, 1482 rpm;1482;37;0;0"}

    Public Shared coupl() As String =
     {"Diameter;cost,percentage na korting",
      "58 mm, n-eupexB;102.7;0.55",
      "68 mm, n-eupexB;111.24;0.55",
      "80 mm, n-eupexB;127.5;0.55",
      "95 mm, n-eupexB;159.58;0.55",
      "160 mm, n-eupexB;294.50;0.45",
      "flender, n-eupexB180;264.94;1",
      "flender, n-eupexB225;313.77;1",
      "N-eupexA250;509.57;1",
      "flender, n-eupexB250;387.75;1",
      "flender, n-eupexA280;763.26;1",
      "flender, n-eupexA315;852.50;1",
      "flender, n-eupexA400;1545.50;1",
      "200 mm, n-eupexB;480;0.45",
      "110 mm, n-eupexA;291.6;0.55",
      "125 mm, n-eupexA;407.35;0.55",
      "140 mm, n-eupexA;534.4;0.55",
      "160 mm, n-eupexA;742.05;0.55",
      "180 mm, n-eupexA;904.4;0.55",
      "200 mm, n-eupexA;1173.5;0.55",
      "225 mm, n-eupexA;1580;0.55",
      "250 mm, n-eupexA;1907.1;0.55",
      "280 mm, n-eupexA;2339;0.55",
      "315 mm, n-eupexA;3190.7;0.55",
      "350 mm, n-eupexA;4433;0.55",
      "400 mm, n-eupexA;5778;0.55",
      "440 mm, n-eupexA;7167.5;0.55",
      "480 mm, n-eupexA;8983.5;0.55"}


    Public Shared ppaint() As String =
     {"Description;cost",
      "10-20m2 75um zink compound;13.25",
      "20-100m2 75um zink compound;12.50",
      "10-20m2 150um primer en epoxy (binnen);17.0",
      "20-100m2 150um primer en epoxy (binnen);16.40",
      "10-20m2 150um primer en polyurethaan (buiten);17.90",
      "20-100m2 150um primer en polyurethaan (buiten);17.25",
      "10-20m2 225um primer, midcoat, polyurethaan (buiten);18.60",
      "20-100m2 225um primer, midcoat, polyurethaan (buiten);18.15",
      "10-20m2 330um primer, midcoat, polyurethaan (buiten);20.75",
      "20-100m2 330um primer, midcoat, polyurethaan (buiten);20.0",
      "10-20m2 75um primer, zincsilicaat -90C/+400C;13.25",
      "20-100m2 75um primer,  zincsilicaat -90C tot +400C;12.50",
      "10-20m2 120um primer,  zincsilicaat -90C tot +400C;19.50",
      "20-100m2 120um primer,  zincsilicaat -90C tot +400C;18.0",
      "10-20m2 250um primer, midcoat, polyurethaan;23.0",
      "20-100m2 250um primer, midcoat, polyurethaan;22.0"}


    Public Shared lager() As String = 'T=trekbus, C=cylindrisch, zie SKFboekje
     {"diameter;prijs",
     "40 mm Trekbus;120.45",
     "50 mm Trekbus;152.57",
      "55 mm Trekbus;170",
      "60 mm Trekbus;196.23",
      "65 mm Trekbus;250",
      "70 mm Trekbus;311.98",
      "75 mm Trekbus;315",
      "80 mm Trekbus;318.59",
      "85 mm Trekbus;370",
      "90 mm Trekbus;420.29",
      "100 mm Trekbus;553.54",
      "110 mm Trekbus;590.63",
      "115 mm Trekbus;729.07",
      "120 mm Trekbus;800",
      "125 mm Trekbus;901.85",
      "130 mm Trekbus;1000",
      "135 mm Trekbus;1111.90",
      "140 mm Trekbus;1385.19",
      "180 mm Trekbus;2505.50",
      "40 mm Cylindrisch;94.39",
      "50 mm Cylindrisch;106.76",
      "55 mm Cylindrisch;120",
      "60 mm Cylindrisch;142.49",
      "65 mm Cylindrisch;170",
      "70 mm Cylindrisch;260.12",
      "75 mm Cylindrisch;270",
      "80 mm Cylindrisch;280.69",
      "85 mm Cylindrisch;290",
      "90 mm Cylindrisch;298.29",
      "95 mm Cylindrisch;340",
      "100 mm Cylindrisch;395.47",
      "110 mm Cylindrisch;526.52",
      "120 mm Cylindrisch;629.66",
      "130 mm Cylindrisch;777.68",
      "140 mm Cylindrisch;962.66",
      "150 mm Cylindrisch;1187.03",
      "160 mm Cylindrisch;1474.77",
      "210 mm Cylindrisch;1700",
      "360 mm Cylindrisch;3500"}

    Public Shared astap_dia() As String =   'tbv as diameter selectie
      {"Dia;empty",
      "100 ;0",
      "110 ;0",
      "120 ;0",
      "140 ;0",
      "160 ;0",
      "200 ;0",
      "210 ;0",
      "260 ;0",
      "310 ;0",
      "350 ;0",
      "400 ;0",
      "500 ;0"}

    Public Shared Flight_dia() As String =   'tbv screw diameter selectie
      {"Flight_dia;empty",
      "280 ;0",
      "330 ;0",
      "400 ;0",
      "500 ;0",
      "630 ;0",
      "800 ;0",
      "1000 ;0",
      "1200 ;0",
      "1400 ;0"}

    Public Shared pakking() As String =
     {"Merk;maat;prijs",
      "Flowtite wit, 3*1.5  ;53",
      "Flowtite wit, 5*2    ;73",
      "Flowtite wit, 7*2.5  ;83",
      "Flowtite wit, 10*3   ;103",
      "230x133x6 Silicone wit;8.15"}


    Public Shared emotor() As String = {"3.0; 1500", "4.0; 1500", "5.5; 1500", "7.5; 1500", "11;  1500", "15; 1500", "22; 1500",
                                       "30  ; 1500", "37;  1500", "45;  1500", "55;  1500", "75; 1500", "90; 1500",
                                       "110 ; 1500", "132; 1500", "160; 1500", "200; 1500"}


    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Fan_sizing_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Fan_rapport_copy\"
    Dim dirpath_Home As String = "C:\Temp\"

    Public Shared _diam_flight As Double                         '[m]
    Public Shared _pipe_OD, _pipe_ID, _pipe_wall As Double
    Public Shared pipe_Ix, pipe_Wx, pipe_Wp As Double            'Lineair en polair weerstand moment
    Public Shared pitch As Double
    Public Shared installed_power As Double
    Public Shared sigma02, sigma_fatique, Elast As Double
    Public Shared inlet_length, conv_length, product_density As Double
    Public Shared _angle As Double
    Public Shared speed As Double
    Public Shared flow_hr As Double
    Public Shared density As Double
    Public Shared filling_perc As Double
    Public Shared progress_resistance As Double                   'Friction from product to steel

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")


        For hh = 0 To (UBound(_inputs) - 1)              'Fill combobox1
            words = _inputs(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 225                   'Grafite ore

        '-------Fill combobox2, Steel selection------------------
        For hh = 0 To (UBound(steel) - 1)               'Fill combobox 2 with steel data
            words = steel(hh).Split(CType(";", Char()))
            ComboBox2.Items.Add(words(0))
        Next hh
        ComboBox2.SelectedIndex = 8                     'rvs 304


        '-------Fill combobox5, emotor selection------------------
        For hh = 0 To (UBound(emotor) - 1)               'Fill combobox 5 emotor data
            words = emotor(hh).Split(CType(";", Char()))
            ComboBox5.Items.Add(words(0))
        Next hh
        ComboBox5.SelectedIndex = 0

        Screw_combo_init()
        Pipe_dia_combo_init()
        Pipe_wall_combo_init()
        Motorreductor()
        Coupling_combo()
        Lager_combo()
        Astap_combo()

        Paint_combo()
        Pakking_combo()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, Button1.Click, TabPage1.Enter, ComboBox11.SelectedValueChanged
        Calculate()
        Calulate_stress_1()
    End Sub

    Private Sub Calculate()
        Dim cap_hr As Double    '100% Çapacity conveyor [m3/hr]
        Dim cap_act As Double   'actual Çapacity conveyor [m3/hr]
        Dim iso_forward As Double       'Power for forward motion
        Dim iso_incline As Double       'Power inclination
        Dim iso_no_product As Double    'Power for seals + bearings
        Dim iso_power As Double         'Total Power 
        Dim height As Double            'Height difference due to inclination 
        Dim mekog As Double             'Mekog installed power
        Dim flight_speed As Double      'Flight speed
        Dim r_time As Double            'Flight speed

        '-------------- get data----------
        Double.TryParse(CType(ComboBox11.SelectedItem, String), _diam_flight)
        _diam_flight /= 1000                                    '[m] -> [mm]
        TextBox18.Text = CType(_diam_flight * CDbl(1000.ToString), String)
        TextBox16.Text = CType(_pipe_OD * CDbl(1000.ToString), String)                'Pipe diameter [m]

        pitch = _diam_flight * NumericUpDown2.Value              '[-]
        conv_length = NumericUpDown3.Value                       'Conveyor length [m]
        TextBox19.Text = conv_length.ToString

        _angle = NumericUpDown4.Value                            '[degree]
        speed = NumericUpDown7.Value                             '[rpm]

        flight_speed = speed / 60 * PI * _diam_flight
        TextBox11.Text = Round(flight_speed, 2).ToString 'Flight speed [m/s]

        Label135.Visible = CBool(IIf(flight_speed > 1.0, True, False))

        If speed > 45 Then
            NumericUpDown7.BackColor = Color.Red
        Else
            NumericUpDown7.BackColor = Color.Yellow
        End If

        flow_hr = NumericUpDown5.Value * 1000           '[kg/hr]
        density = NumericUpDown6.Value                  '[kg/m3]
        progress_resistance = NumericUpDown9.Value      '[-]

        '--------------- now calc-----------------

        cap_hr = PI / 4 * (_diam_flight ^ 2 - _pipe_OD ^ 2) * pitch * speed * 60          ' [m]
        cap_hr = cap_hr * (100 - _angle * 2) / 100                                         ' capacity loss due to inclination (2% per degree)

        cap_act = flow_hr / density
        filling_perc = Round(cap_act / cap_hr * 100, 1)

        If filling_perc > 100 Then filling_perc = 100

        TextBox1.BackColor = CType(IIf(filling_perc > 40, Color.Red, Color.LightGreen), Color)


        '--------------- ISO 7119 -----------------
        height = conv_length * Sin(_angle / 360 * 2 * PI)

        iso_forward = flow_hr * conv_length * 9.91 * progress_resistance / (3600 * 1000)    'Forwards [kW]
        iso_incline = flow_hr * height * 9.81 / (3600 * 1000)                               'Uphill [kW]
        iso_no_product = _diam_flight * conv_length / 20                                     'Power for seals 0. + bearings [kW]

        iso_power = Round(iso_forward + iso_incline + iso_no_product, 1)

        '--------------- MEKOG -----------------
        mekog = Round(flow_hr * conv_length / (40 * 1.36 * 1000), 1)    '[kW]

        '-------------- Retention time --------------------
        r_time = conv_length / (speed / 60 * pitch)                     '[sec]

        '--------------- present results------------
        TextBox1.Text = filling_perc.ToString
        TextBox3.Text = iso_power.ToString
        TextBox4.Text = mekog.ToString
        TextBox110.Text = Round(r_time, 0).ToString

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Save_to_disk()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabControl1.Enter, RadioButton8.CheckedChanged, RadioButton7.CheckedChanged, RadioButton6.CheckedChanged, RadioButton4.CheckedChanged, NumericUpDown35.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown25.ValueChanged, ComboBox9.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox7.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, ComboBox13.SelectedIndexChanged, ComboBox12.SelectedIndexChanged, ComboBox10.SelectedIndexChanged, CheckBox8.CheckedChanged, CheckBox5.CheckedChanged, CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, CheckBox4.CheckedChanged, CheckBox7.CheckedChanged, CheckBox6.CheckedChanged, CheckBox9.CheckedChanged, TabPage4.Enter
        Costing_material()
    End Sub

    'Materiaal in de conveyor
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim words() As String = _inputs(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            NumericUpDown6.Value = CDec(words(5)) 'Density max
            NumericUpDown9.Value = CDec(words(6)) 'Material factor
            Label37.Text = "CEMA material code " & words(1)
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, TabPage5.Enter, NumericUpDown34.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown27.ValueChanged
        Costing_material()
    End Sub
    'Please note complete calculation in [m] nit [mm]
    Private Sub Calulate_stress_1()
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
            words = emotor(ComboBox5.SelectedIndex).Split(CType(";", Char()))
            Double.TryParse(words(0), installed_power)
        End If

        If (ComboBox2.SelectedIndex > -1) Then      'Prevent exceptions
            words = steel(ComboBox2.SelectedIndex).Split(CType(";", Char()))
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
            TextBox7.Text = CType(sigma02, String)
            sigma_fatique = sigma02 * 0.35                   'Fatique stress uitgelegd op oneindige levensduur
            TextBox8.Text = Round(sigma_fatique, 0).ToString
        End If

        If ComboBox3.SelectedIndex > -1 Then
            words = pipe(ComboBox3.SelectedIndex).Split(CType(";", Char()))

            Double.TryParse(words(2), _pipe_OD)
            _pipe_OD /= 1000                             'Outside Diameter [m]
            pipe_OR = _pipe_OD / 2                       'Radius [mm]
            _pipe_wall = CDbl(ComboBox6.SelectedItem)   'Wall thickness [mm]
            _pipe_wall /= 1000
            _pipe_ID = (_pipe_OD - 2 * _pipe_wall)         'Inside diameter [mm]
            pipe_IR = _pipe_ID / 2                       'Inside radius [mm]

            weight_m = PI / 4 * (_pipe_OD ^ 2 - _pipe_ID ^ 2) * 7850          'Weight per meter [kg/m]

            TextBox13.Text = Round(weight_m, 1).ToString                    'gewicht per meter
            TextBox16.Text = Round(_pipe_OD * 1000, 1).ToString              'Diameter [m]

            '---------------- Traagheids moment Ix= PI/64.(D^4-d^4)---------------------
            pipe_Ix = PI / 64 * (_pipe_OD ^ 4 - _pipe_ID ^ 4)                  '[m4]
            TextBox26.Text = Round(pipe_Ix * 1000 ^ 4, 0).ToString

            '---------------- Weerstand moment Buiging  Wx= PI/32.(D^4-d^4)/D---------------------
            pipe_Wx = PI / 32 * (_pipe_OD ^ 4 - _pipe_ID ^ 4) / _pipe_OD        '[m3]
            TextBox14.Text = Round(pipe_Wx * 1000 ^ 3, 0).ToString

            '---------------- Weerstand moment Torsie (polair)  Wp= PI/16.(D^4-d^4)/D --------------
            pipe_Wp = PI / 16 * (_pipe_OD ^ 4 - _pipe_ID ^ 4) / _pipe_OD       '[m3]
            TextBox15.Text = Round(pipe_Wp * 1000 ^ 3, 0).ToString


            '============================================calc load ==============================================================================
            '====================================================================================================================================

            '---------------- gewicht flight---mm dik----------------------------------
            flight_hoogte = (_diam_flight - _pipe_OD / 1000) / 2                                  '[m]
            flight_lengte_buiten = Sqrt((PI * _diam_flight) ^ 2 + (pitch) ^ 2)

            flight_lengte_binnen = Sqrt((PI * _pipe_OD / 1000) ^ 2 + (pitch) ^ 2)
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
            Radius_transport = (_diam_flight + _pipe_OD) / 4                  'Acc Jos (D+d)/4
            F_tangent = P_torque / Radius_transport
            Q_load_2 = F_tangent / conv_length                              'Transport kracht geeft doorbuiging pijp
            Q_load_3 = _pipe_OD * kolom_height * product_density * 9.91      'gelijkmatige belasting op de pijp door materiaal kolom
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
            'Elast = 210 * 1000 ^ 2                                   '[N/mm2]  ??????
            Elast = NumericUpDown1.Value * 1000 '[N/mm2]
            Q_Deflect_max = (5 * Q_load_comb / 1000 * conv_length ^ 4) / (384 * Elast * pipe_Ix)
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown11.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown13.ValueChanged, TabPage2.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, CheckBox1.CheckedChanged, ComboBox3.SelectedIndexChanged
        Calulate_stress_1()
    End Sub
    Private Sub Screw_dia_combo()
        Dim words() As String

        If (ComboBox11.SelectedIndex > -1) Then      'Prevent exceptions
            words = Flight_dia(ComboBox11.SelectedIndex).Split(CType(";", Char()))
            Double.TryParse(words(0), _diam_flight)
            _diam_flight /= 1000                    'Trough width[m]

            TextBox18.Text = Round(_diam_flight * 1000, 3).ToString
            MessageBox.Show(_diam_flight.ToString)
        End If
    End Sub

    Private Sub Pipe_dia_combo_init()
        Dim words() As String

        ComboBox3.Items.Clear()

        '-------Fill combobox3, Pipe selection------------------
        For hh = 0 To (UBound(pipe) - 1)                'Fill combobox 3 with pipe data
            words = pipe(hh).Split(CType(";", Char()))
            ComboBox3.Items.Add(Trim(words(2)))
            ComboBox9.Items.Add(Trim(words(2)))
        Next hh
        ComboBox3.SelectedIndex = 2
        ComboBox9.SelectedIndex = 2

        words = pipe(ComboBox3.SelectedIndex).Split(CType(";", Char()))
        Double.TryParse(words(2), _pipe_OD)
        _pipe_OD /= 1000                                         'Outside Diameter [m]
        TextBox16.Text = Round(_pipe_OD * 1000, 1).ToString      'Diameter [mm]
    End Sub
    Private Sub Pipe_wall_combo_init()
        Dim words() As String
        Dim temp As Double

        ComboBox6.Items.Clear()
        '-------Fill combobox6, pipe wall selection------------------
        words = pipe(ComboBox3.SelectedIndex).Split(CType(";", Char()))  'Fill combobox 6 pipe wall data
        For hh = 3 To 5
            If Double.TryParse(words(hh), temp) Then
                ComboBox6.Items.Add(Trim(words(hh)))
            End If
        Next
        ComboBox6.SelectedIndex = 1
    End Sub
    Private Sub Motorreductor()
        Dim words() As String

        ComboBox4.Items.Clear()
        '-------Fill combobox4,  selection------------------
        For hh = 1 To (UBound(motorred))                'Fill combobox 3 with pipe data
            words = motorred(hh).Split(CType(";", Char()))
            ComboBox4.Items.Add(Trim(words(0)))
        Next hh
        ComboBox4.SelectedIndex = 2
    End Sub
    Private Sub Coupling_combo()
        Dim words() As String

        ComboBox7.Items.Clear()
        '-------Fill combobox7,  selection------------------
        For hh = 1 To (UBound(coupl))                'Fill combobox 7 with coupling data
            words = coupl(hh).Split(CType(";", Char()))
            ComboBox7.Items.Add(Trim(words(0)))
        Next hh
        ComboBox7.SelectedIndex = 1
    End Sub
    Private Sub Lager_combo()
        Dim words() As String

        ComboBox8.Items.Clear()
        '-------Fill combobox8,  selection------------------
        For hh = 1 To lager.Length - 1                'Fill combobox 8 with lager data
            words = lager(hh).Split(CType(";", Char()))
            ComboBox8.Items.Add(words(0))
        Next hh
        ComboBox8.SelectedIndex = 1
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer

        'Start Word and open the document template. 
        font_sizze = 9
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = font_sizze + 3
        oPara1.Range.Font.Bold = CInt(True)
        oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = font_sizze + 1
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = CInt(False)
        oPara2.Range.Text = "Kostencalculatie van schroeftransporteur" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Project Name"
        oTable.Cell(row, 2).Range.Text = TextBox23.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project number "
        oTable.Cell(row, 2).Range.Text = TextBox35.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Schroefnummer "
        oTable.Cell(row, 2).Range.Text = TextBox53.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Author "
        oTable.Cell(row, 2).Range.Text = Environment.UserName
        row += 1
        oTable.Cell(row, 1).Range.Text = "Date "
        oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 14, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input Data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter Flight"
        oTable.Cell(row, 3).Range.Text = ComboBox11.Text
        oTable.Cell(row, 2).Range.Text = "[mm]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter pipe"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox9.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox45.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Wall thickness pipe"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox6.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Pitch"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown2.Value, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Blad dikte"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown8.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox46.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Toerental"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown7.Value, String)
        oTable.Cell(row, 2).Range.Text = "[rpm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Installed Power"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox5.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[kW]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Conveyor length"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown3.Value, String)
        oTable.Cell(row, 2).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Inclination angle"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown4.Value, String)
        oTable.Cell(row, 2).Range.Text = "[deg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Staalsoort"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox2.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Temperature"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown11.Value, String)
        oTable.Cell(row, 2).Range.Text = "[C]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Product type"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox1.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        '---- -----
        oTable.Cell(row, 1).Range.Text = "Capacity"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown5.Value, String)
        oTable.Cell(row, 2).Range.Text = "[ton/hr]"


        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(1.6)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(0.6)

        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 12, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Motorreductor"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox4.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Coupling"
        NewMethod(oTable, row)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Lager asdiameter"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox8.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Aantal certificaten "
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown27.Value, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dikte kop-en staartplaat"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown10.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox42.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dikte schroefblad"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown8.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox46.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dikte trog"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown14.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox47.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dikte deksel"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown15.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox48.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Astap diameter"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox13.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox54.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"

        row += 1
        oTable.Cell(row, 5).Range.Text = "_____"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Total sheet steel"
        oTable.Cell(row, 5).Range.Text = TextBox109.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(1.6)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(0.6)

        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 11, 8)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Kosten"
        row += 1
        oTable.Rows.Item(2).Range.Font.Bold = CInt(True)
        oTable.Cell(row, 6).Range.Text = "Material"
        oTable.Cell(row, 1).Range.Text = "Labour"
        'row += 1
        'oTable.Cell(row, 3).Range.Text = "Hours"
        'oTable.Cell(row, 5).Range.Text = "Costs"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Engineering"
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown30.Value, String)
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 5).Range.Text = TextBox55.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown33.Value, String)
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 5).Range.Text = TextBox70.Text

        row += 1
        oTable.Cell(row, 1).Range.Text = "Work shop"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown34.Value, String)
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 5).Range.Text = TextBox72.Text
        oTable.Cell(row, 4).Range.Text = "[€]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Total hours"
        oTable.Cell(row, 3).Range.Text = TextBox106.Text
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 5).Range.Text = TextBox98.Text
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 6).Range.Text = "Total cost material"
        oTable.Cell(row, 8).Range.Text = TextBox103.Text
        oTable.Cell(row, 7).Range.Text = "[€]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Percentage labour "
        oTable.Cell(row, 5).Range.Text = TextBox101.Text
        oTable.Cell(row, 4).Range.Text = "[%]"
        oTable.Cell(row, 6).Range.Text = "Percentage material"
        oTable.Cell(row, 8).Range.Text = TextBox100.Text
        oTable.Cell(row, 7).Range.Text = "[%]"

        row += 1

        oTable.Cell(row, 1).Range.Text = "Total cost price"
        oTable.Cell(row, 5).Range.Text = TextBox73.Text
        oTable.Cell(row, 4).Range.Text = "[€]"

        row += 1
        oTable.Rows.Item(11).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(11).Range.Font.Size = font_sizze + 1
        oTable.Cell(row, 1).Range.Text = "Total sale price"
        oTable.Cell(row, 5).Range.Text = TextBox75.Text
        oTable.Cell(row, 4).Range.Text = "[€]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(1.2)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(1.4)
        oTable.Columns.Item(6).Width = oWord.InchesToPoints(1.5)
        oTable.Columns.Item(7).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(8).Width = oWord.InchesToPoints(0.6)
    End Sub

    Private Sub NewMethod(oTable As Word.Table, row As Integer)
        oTable.Cell(row, 3).Range.Text = CType(ComboBox7.SelectedItem, String)
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs)
        Screw_dia_combo()
        Calculate()
    End Sub

    Private Sub Astap_combo()
        Dim words() As String

        ComboBox13.Items.Clear()
        '-------Fill combobox------------------
        For hh = 1 To astap_dia.Length - 1                'Fill combobox 3 with pipe data
            words = astap_dia(hh).Split(CType(";", Char()))
            ComboBox13.Items.Add(words(0))
        Next hh
        ComboBox13.SelectedIndex = 1
    End Sub
    Private Sub Screw_combo_init()
        Dim words() As String

        ComboBox11.Items.Clear()
        '-------Fill combobox------------------
        For hh = 1 To Flight_dia.Length - 1                'Fill combobox 11 with flight data
            words = Flight_dia(hh).Split(CType(";", Char()))
            ComboBox11.Items.Add(words(0))
        Next hh
        ComboBox11.SelectedIndex = 2
    End Sub

    Private Sub Paint_combo()
        Dim words() As String

        ComboBox12.Items.Clear()
        '-------Fill combobox ------------------
        For hh = 1 To ppaint.Length - 1                'Fill combobox 3 with pipe data
            words = ppaint(hh).Split(CType(";", Char()))
            ComboBox12.Items.Add(words(0))
        Next hh
        ComboBox12.SelectedIndex = 1
    End Sub
    Private Sub Pakking_combo()
        Dim words() As String

        ComboBox10.Items.Clear()
        '-------Fill combobox-----------------
        For hh = 1 To pakking.Length - 1                'Fill combobox 3 with pipe data
            words = pakking(hh).Split(CType(";", Char()))
            ComboBox10.Items.Add(words(0))
        Next hh
        ComboBox10.SelectedIndex = 3
    End Sub

    'Save data and line chart to file
    Private Sub Save_to_disk()
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
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Print_word()
    End Sub

    Private Sub Print_word()
        Dim oWord As Word.Application ' = Nothing
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename, str As String

        Try
            oWord = New Word.Application()

            'Start Word and open the document template. 
            font_sizze = 9
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = font_sizze + 3
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Font.Size = font_sizze + 1
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "Screw Conveyor Stress calculation " & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox65.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Project number "
            oTable.Cell(row, 2).Range.Text = TextBox66.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Machine Id "
            oTable.Cell(row, 2).Range.Text = TextBox67.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author "
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date "
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a 16 x 3 table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 18, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Conveyor Data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Diameter trough"
            oTable.Cell(row, 2).Range.Text = ComboBox11.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Diameter pipe"
            oTable.Cell(row, 2).Range.Text = ComboBox3.Text & " x " & ComboBox6.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Pitch"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown2.Value, String)
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Blade thicknes"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown8.Value, String)
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Length"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown3.Value, String)
            oTable.Cell(row, 3).Range.Text = "[m]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inclination angle"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown4.Value, String)
            oTable.Cell(row, 3).Range.Text = "[degree]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Speed"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown7.Value, String)
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flight tip speed"
            oTable.Cell(row, 2).Range.Text = TextBox11.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Filling "
            oTable.Cell(row, 2).Range.Text = TextBox1.Text
            oTable.Cell(row, 3).Range.Text = "[%]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Power ISO 7119"
            oTable.Cell(row, 2).Range.Text = TextBox3.Text
            oTable.Cell(row, 3).Range.Text = "[kW]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Power MEKOG"
            oTable.Cell(row, 2).Range.Text = TextBox4.Text
            oTable.Cell(row, 3).Range.Text = "[kW]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Power Installed"
            oTable.Cell(row, 2).Range.Text = ComboBox5.Text
            oTable.Cell(row, 3).Range.Text = "[kW]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Product"
            str = ComboBox1.Text
            If Len(str) > 22 Then str = str.Substring(0, 22)
            oTable.Cell(row, 2).Range.Text = str

            row += 1
            oTable.Cell(row, 1).Range.Text = "Product Flow"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown5.Value, String)
            oTable.Cell(row, 3).Range.Text = "[ton/hr]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Product Density"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown6.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Forward resistance"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown9.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
            oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.5)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------- Results----------------------
            'Insert a 5 x 3 table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Calculation Results"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bending Stress"
            oTable.Cell(row, 2).Range.Text = TextBox9.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm^2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Torque Stress"
            oTable.Cell(row, 2).Range.Text = TextBox12.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm^2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Combined Stress"
            oTable.Cell(row, 2).Range.Text = TextBox21.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm^2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Selected steel"
            oTable.Cell(row, 2).Range.Text = ComboBox2.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max. Fatique stress"
            oTable.Cell(row, 2).Range.Text = TextBox8.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Center pipe load (weight only)"
            oTable.Cell(row, 2).Range.Text = TextBox22.Text
            oTable.Cell(row, 3).Range.Text = "[N/m]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max Flex"
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
            oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.5)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            ''-------------- Checks-------
            'Insert a 5 x 1 table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 1)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Checks "
            row += 1
            If (TextBox11.BackColor = Color.Red) Then
                oTable.Cell(row, 1).Range.Text = "NOK, for ATEX, Flight speed > 1 m/s"
            Else
                oTable.Cell(row, 1).Range.Text = "OK, Flight speed for ATEX applications"
            End If
            row += 1
            If NumericUpDown7.BackColor = Color.Red Then
                oTable.Cell(row, 1).Range.Text = "NOK, Rotational speed > 45 rpm, too fast"
            Else
                oTable.Cell(row, 1).Range.Text = "OK, Rotational speed"
            End If
            row += 1
            If TextBox1.BackColor = Color.Red Then
                oTable.Cell(row, 1).Range.Text = "NOK, Filling percentage > 40%"
            Else
                oTable.Cell(row, 1).Range.Text = "OK, Filling percentage"
            End If
            row += 1
            If TextBox21.BackColor = Color.Red Then
                oTable.Cell(row, 1).Range.Text = "NOK, Combined pipe stress too high"
            Else
                oTable.Cell(row, 1).Range.Text = "OK, Combined pipe stress"
            End If
            oTable.Columns(1).Width = oWord.InchesToPoints(4.0)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '--------------Save file word file------------------
            'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx

            ufilename = "Conveyor_report_" & DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss") & ".docx"

            If Directory.Exists(dirpath_Rap) Then
                ufilename = dirpath_Rap & ufilename
            Else
                ufilename = dirpath_Home & ufilename
            End If
            oWord.ActiveDocument.SaveAs(ufilename.ToString)

        Catch ex As Exception
            MessageBox.Show("Line 683, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Costing_material()
        Dim rho_materiaal, rho_kunststof, conv_length As Double
        Dim dikte_trog, opp_trog, kopstaartplaat, weight_kopstaartplaat, kg_trog As Double
        Dim weight_pipe, dikte_deksel, speling_trog, diam_schroef As Double
        Dim kg_inlaat, kg_uitlaat, kg_lining, dikte_lining As Double
        Dim kg_deksel, kg_voet, kg_schroefblad, total_kg_plaat As Double
        Dim hoek_spoed As Double
        Dim nr_flights, spoed As Double
        Dim kg_astap, dia_astap, lengte_astap As Double
        Dim kg_afschermkap As Double
        Dim tot_opperv_paint As Double
        Dim oppb_afschermkap, oppb_astap, oppb_voet, oppb_uitlaat, oppb_inlaat As Double
        Dim oppb_deksel, oppb_trog, oppb_kopstaartplaat, oppb_schroefblad, oppb_pipe As Double
        Dim cost_kopstaartplaat, cost_trog, cost_pipe, cost_deksel, cost_inlaat, cost_uitlaat As Double
        Dim cost_voet, cost_schroefblad, cost_astap, cost_lining, cost_afschermkap As Double
        Dim cost_paint, cost_painting, cost_cutting As Double
        Dim cost_motorreductor, cost_koppeling, cost_lagers, cost_stopbuspakking As Double
        Dim cost_pakking, cost_hang, cost_transport As Double
        Dim cost_stopbus As Double
        Dim certificate_cost, totalplate_cost, total_cost As Double
        Dim uren_engineering, uren_project, uren_fabrieks, tot_uren As Double
        Dim engineering_prijs_uur, project_prijs_uur, fabriek_prijs_uur As Double
        Dim prijs_engineering, prijs_project, prijs_fabriek As Double
        Dim tot_prijsarbeid, geheel_totprijs, dekking, marge_cost, verkoopprijs, perc_mater, perc_arbeid As Double

        conv_length = NumericUpDown3.Value             'lengte van de trog [m]
        TextBox40.Text = ComboBox2.Text                'materiaalsoort staal
        TextBox41.Text = (_pipe_OD * 1000).ToString    'diameter pijp
        TextBox51.Text = CType(NumericUpDown3.Value, String)          'lengte trog
        TextBox52.Text = ComboBox5.Text                'vermogen aandrijving
        TextBox44.Text = _diam_flight.ToString         'diameter flight

        '---------------------------------------------- PRICES -----------------------------------------
        '-----------------------------------------------------------------------------------------------
        TextBox84.Text = "3.25"                 'lining [€/kg]
        TextBox85.Text = "10.00"                'alu afschermkap [€/kg]
        TextBox113.Text = "0.50"                'snijkosten [€/kg]

        Select Case True
            Case (RadioButton6.Checked)         'staal, s235JR
                rho_materiaal = 7850
                TextBox93.Text = "0.88"         'kop staart  [€/kg]
                TextBox94.Text = "2.09"         'schroefpijp
                TextBox95.Text = "3.03"         'schroefblad
                TextBox96.Text = "0.78"         'trog
                TextBox97.Text = "0.78"         'deksel
                TextBox92.Text = "2.09"         'astap ronde staf afm 60
            Case (RadioButton7.Checked)         'rvs304, warmgewalst
                rho_materiaal = 8000
                TextBox93.Text = "2.45"         'kop staart 
                TextBox94.Text = "2.45"         'schroefpijp
                TextBox95.Text = "2.45"         'schroefblad
                TextBox96.Text = "2.45"         'trog
                TextBox97.Text = "2.45"         'deksel
                TextBox92.Text = "1.52"         'astap [€/kg] materiaal is standaard van staal
            Case (RadioButton8.Checked)         'rvs316, warmgewalst(zie vtke-151401)
                rho_materiaal = 7860
                TextBox93.Text = "4.07"         'kop staart 
                TextBox94.Text = "7.57"         'schroefpijp
                TextBox95.Text = "6.07"         'schroefblad
                TextBox96.Text = "4.07"         'trog
                TextBox97.Text = "4.07"         'deksel
                TextBox92.Text = "2.09"         'astap [€/kg] materiaal is standaard van staal
        End Select

        Try
            Dim words1() As String = lager(ComboBox8.SelectedIndex + 1).Split(CType(";", Char()))
            cost_lagers = CDbl(words1(1))

            Dim words2() As String = coupl(ComboBox7.SelectedIndex + 1).Split(CType(";", Char()))
            cost_koppeling = CDbl(words2(1)) * CDbl(words2(2))                                         'inclusief kortingspercentage van 45%
            If Not CheckBox3.Checked Then cost_koppeling = 0

            Dim words3() As String = motorred(ComboBox4.SelectedIndex + 1).Split(CType(";", Char()))
            cost_motorreductor = CDbl(words3(3))
            If Not CheckBox2.Checked Then cost_motorreductor = 0

            Dim words4() As String = ppaint(ComboBox12.SelectedIndex + 1).Split(CType(";", Char()))
            cost_paint = CDbl(words4(1))
            If Not CheckBox6.Checked Then cost_paint = 0

            Dim words5() As String = pakking(ComboBox10.SelectedIndex + 1).Split(CType(";", Char()))
            cost_pakking = CDbl(words5(1))

            cost_inlaat = 300   'inlaat chute
            cost_uitlaat = 300  'Uitlaat chute
            cost_voet = 100     'Conveyor supports

        Catch ex As Exception
            'MessageBox.Show(ex.Message & "Line 1778")  ' Show the exception's message.
        End Try


        '----------------------------------------WEIGHT + AREA CALCULATIONS-----------------------------------------
        '----------------------------------------------------------------------------------------------
        dikte_trog = NumericUpDown14.Value / 1000


        Select Case True                'Pijpschroef oppervlak
            Case (RadioButton4.Checked)
                opp_trog = 2 * PI * ((_diam_flight / 2) * dikte_trog)
                kopstaartplaat = _diam_flight ^ 2
            Case (RadioButton5.Checked) 'Trogschroef oppervlak
                opp_trog = (PI * (_diam_flight / 2) * dikte_trog + 2 * dikte_trog * (0.045 + _diam_flight / 2) + 0.075 * dikte_trog)   'troghoogte=trogbreedte/2+45mm, flens= 0.05+0.025
                kopstaartplaat = (_diam_flight * (_diam_flight + 0.045))
        End Select

        weight_kopstaartplaat = kopstaartplaat * (NumericUpDown10.Value / 1000) * rho_materiaal
        oppb_kopstaartplaat = 2 * kopstaartplaat

        kg_trog = 2 * weight_kopstaartplaat + opp_trog * conv_length * rho_materiaal
        oppb_trog = 2 * kopstaartplaat + 2 * opp_trog * conv_length / dikte_trog                'kuip zowel uitwendig als inwendig

        Double.TryParse(CType(ComboBox9.SelectedItem, String), _pipe_OD)         ' ComboBox3 = ComboBox9
        _pipe_OD = _pipe_OD / 1000
        _pipe_wall = CDbl(ComboBox6.SelectedItem)
        _pipe_wall /= 1000
        _pipe_ID = (_pipe_OD - 2 * _pipe_wall)
        weight_pipe = rho_materiaal * PI / 4 * (_pipe_OD ^ 2 - _pipe_ID ^ 2) * conv_length
        oppb_pipe = _pipe_OD * PI * conv_length

        If _diam_flight > 0.3015 Then                          'in [m], radiale speling schroef in kuip: tot diam 0.3m 7.5 mm, daarboven 10mm
            speling_trog = 0.01
        Else
            speling_trog = 0.0075
        End If
        diam_schroef = _diam_flight - 2 * speling_trog

        dikte_deksel = NumericUpDown15.Value / 1000
        kg_deksel = conv_length * dikte_deksel * (_diam_flight + 0.075) * rho_materiaal     '50mm voor de horizontale flens en 25mm voor het stukje naar beneden
        oppb_deksel = 2 * conv_length * (_diam_flight + 0.075)                              'zowel inwendig als uitwendig



        NumericUpDown12.Value = NumericUpDown8.Value                    'Dikte schroefblad bij tab1 opgegeven
        spoed = diam_schroef * NumericUpDown2.Value
        nr_flights = conv_length / spoed
        hoek_spoed = Atan(spoed / (PI * diam_schroef))                  '[rad]    

        kg_schroefblad = PI * rho_materiaal * (NumericUpDown12.Value / 1000) * 0.25 * nr_flights * (diam_schroef ^ 2 - _pipe_OD ^ 2) / Cos(hoek_spoed)         ' DIT IS DE ECHTE FORMULE!!!!!
        oppb_schroefblad = 2 * (kg_schroefblad / (NumericUpDown12.Value * rho_materiaal / 1000))

        Double.TryParse(CType(ComboBox13.SelectedItem, String), dia_astap)             '[mm] 
        dia_astap = dia_astap / 1000                                    '[m]
        lengte_astap = 1.0                                              'lengte in meters average 1m
        kg_astap = 7850 * lengte_astap * PI * (dia_astap / 2) ^ 2       'het standaardmateriaal is staal, dit is het totale inkoopmateriaal, wat daarna nog wordt gefreesd/gedraaid
        oppb_astap = PI * dia_astap * lengte_astap

        rho_kunststof = 970                                             '[kg/m3] dichtheid HDPE
        dikte_lining = NumericUpDown25.Value / 1000
        kg_lining = rho_kunststof * (PI * _diam_flight + 0.5 * (0.045 + _diam_flight / 2)) * dikte_lining * conv_length

        '---------- estimated weights---------------
        kg_inlaat = 10              '[kg] inlaat chute
        kg_uitlaat = 10             '[kg] uitlaat chute
        kg_voet = 5                 '[kg] conveyor supports
        kg_afschermkap = 10         '[kg] koppelingkap

        '---------- estimated area's---------------
        oppb_inlaat = 1             '[m2]
        oppb_uitlaat = 1            '[m2]
        oppb_voet = 0.5             '[m2]
        oppb_afschermkap = 1        '[m2]

        total_kg_plaat = 2 * weight_kopstaartplaat + kg_trog + kg_inlaat + kg_uitlaat + kg_voet + kg_afschermkap    'Onderdelen van plaat die gesneden worden
        tot_opperv_paint = oppb_voet + oppb_uitlaat + oppb_inlaat + oppb_trog        'Buiten Oppervlak paint onderdelen 


        '----------------------------------------COST CALCULATION-----------------------------------------
        '----------------------------------------------------------------------------------------------
        Try
            cost_painting = cost_paint * tot_opperv_paint
            cost_kopstaartplaat = weight_kopstaartplaat * Double.Parse(TextBox93.Text)
            cost_trog = kg_trog * Double.Parse(TextBox96.Text)
            cost_pipe = weight_pipe * Double.Parse(TextBox94.Text)
            cost_deksel = kg_deksel * Double.Parse(TextBox97.Text)
            If Not CheckBox5.Checked Then cost_deksel = 0

            cost_cutting = total_kg_plaat * Double.Parse(TextBox113.Text)
            If Not CheckBox9.Checked Then cost_cutting = 0

            cost_inlaat *= NumericUpDown20.Value
            cost_uitlaat *= NumericUpDown21.Value
            cost_voet *= NumericUpDown23.Value
            cost_schroefblad = kg_schroefblad * Double.Parse(TextBox95.Text)
            cost_astap = kg_astap * Double.Parse(TextBox92.Text)
            cost_lining = kg_lining * Double.Parse(TextBox84.Text)
            If Not CheckBox4.Checked Then cost_lining = 0
            cost_afschermkap = kg_afschermkap * Double.Parse(TextBox85.Text)
            cost_stopbus = 500                                  '€, te ingewikkeld om precieze prijs te bepalen (2 stuks) 
            If Not CheckBox8.Checked Then cost_stopbus = 0

            cost_hang = NumericUpDown35.Value * 500
            cost_transport = 800                                '€ intern transport
            If Not CheckBox7.Checked Then cost_transport = 0

            totalplate_cost = 2 * cost_kopstaartplaat + cost_trog + cost_pipe + cost_inlaat + cost_uitlaat + cost_voet
            totalplate_cost += cost_schroefblad + cost_astap + cost_lining + cost_afschermkap + cost_stopbus + cost_hang + cost_deksel

        Catch ex As Exception
            'MessageBox.Show(ex.Message & "Line 1904")  ' Show the exception's message.
        End Try

        certificate_cost = 50 * NumericUpDown27.Value               'certificaat €50/stuk 
        total_cost = totalplate_cost + cost_motorreductor + cost_koppeling + cost_lagers + cost_stopbuspakking
        total_cost += cost_painting + certificate_cost + cost_pakking + cost_transport + cost_cutting

        TextBox42.Text = Round(weight_kopstaartplaat, 1).ToString
        TextBox47.Text = Round(kg_trog, 1).ToString
        TextBox45.Text = Round(weight_pipe, 1).ToString
        TextBox48.Text = Round(kg_deksel, 1).ToString
        TextBox46.Text = Round(kg_schroefblad, 1).ToString
        TextBox54.Text = Round(kg_astap, 1).ToString
        TextBox77.Text = Round(kg_lining, 1).ToString
        TextBox76.Text = Round(kg_afschermkap, 1).ToString
        TextBox108.Text = Round(tot_opperv_paint, 1).ToString

        TextBox63.Text = Round(cost_lagers, 2).ToString         'Lagers
        TextBox57.Text = Round(cost_motorreductor, 2).ToString  'Drive
        TextBox58.Text = Round(cost_koppeling, 2).ToString      'Coupling
        TextBox107.Text = Round(cost_painting, 2).ToString      'Paint
        TextBox104.Text = Round(cost_pakking, 2).ToString       'Seals
        TextBox43.Text = Round(cost_cutting, 2).ToString        'Plate cutting

        TextBox56.Text = Round(cost_kopstaartplaat, 1).ToString
        TextBox61.Text = Round(cost_trog, 1).ToString
        TextBox59.Text = Round(cost_pipe, 1).ToString
        TextBox62.Text = Round(cost_deksel, 1).ToString
        TextBox79.Text = Round(cost_inlaat, 1).ToString
        TextBox80.Text = Round(cost_uitlaat, 1).ToString
        TextBox81.Text = Round(cost_voet, 1).ToString
        TextBox60.Text = Round(cost_schroefblad, 1).ToString
        TextBox78.Text = Round(cost_astap, 1).ToString
        TextBox83.Text = Round(cost_lining, 1).ToString
        TextBox82.Text = Round(cost_afschermkap, 1).ToString
        TextBox64.Text = Round(cost_stopbus, 1).ToString
        TextBox88.Text = Round(certificate_cost, 1).ToString
        TextBox102.Text = Round(cost_hang, 1).ToString
        TextBox112.Text = Round(cost_transport, 1).ToString

        ''Tabblad sales price

        uren_engineering = NumericUpDown30.Value
        uren_project = NumericUpDown33.Value
        uren_fabrieks = NumericUpDown34.Value
        engineering_prijs_uur = 80
        project_prijs_uur = 100
        fabriek_prijs_uur = 60
        prijs_engineering = uren_engineering * engineering_prijs_uur
        prijs_project = uren_project * project_prijs_uur
        prijs_fabriek = uren_fabrieks * fabriek_prijs_uur
        tot_uren = uren_engineering + uren_project + uren_fabrieks

        tot_prijsarbeid = prijs_engineering + prijs_project + prijs_fabriek
        geheel_totprijs = total_cost + tot_prijsarbeid
        perc_mater = 100 * total_cost / geheel_totprijs
        perc_arbeid = 100 * tot_prijsarbeid / geheel_totprijs
        dekking = geheel_totprijs * 0.175
        marge_cost = (geheel_totprijs + dekking) * 0.1
        verkoopprijs = geheel_totprijs + dekking + marge_cost

        TextBox109.Text = Round(total_kg_plaat, 0).ToString
        TextBox68.Text = Round(totalplate_cost, 0).ToString
        TextBox105.Text = Round(engineering_prijs_uur, 0).ToString
        TextBox55.Text = Round(prijs_engineering, 1).ToString
        TextBox69.Text = Round(project_prijs_uur, 0).ToString
        TextBox70.Text = Round(prijs_project, 0).ToString
        TextBox71.Text = Round(fabriek_prijs_uur, 0).ToString
        TextBox72.Text = Round(prijs_fabriek, 0).ToString
        TextBox106.Text = Round(tot_uren, 0).ToString

        TextBox88.Text = Round(certificate_cost, 0).ToString
        TextBox111.Text = Round(total_cost, 2).ToString
        TextBox103.Text = Round(total_cost, 1).ToString
        TextBox98.Text = Round(tot_prijsarbeid, 0).ToString
        TextBox100.Text = Round(perc_mater, 0).ToString
        TextBox101.Text = Round(perc_arbeid, 0).ToString
        TextBox73.Text = Round(geheel_totprijs, 0).ToString
        TextBox74.Text = Round(dekking, 0).ToString
        TextBox99.Text = Round(marge_cost, 0).ToString
        TextBox75.Text = Round(verkoopprijs, 0).ToString

    End Sub
End Class

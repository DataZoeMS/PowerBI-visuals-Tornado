import powerbiVisualsApi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { legendInterfaces } from "powerbi-visuals-utils-chartutils";
import { ColorHelper } from "powerbi-visuals-utils-colorutils";
import LegendPosition = legendInterfaces.LegendPosition;

import { TornadoChartSeries } from "./interfaces"

import Card = formattingSettings.SimpleCard;
import CompositeCard = formattingSettings.CompositeCard;
import Model = formattingSettings.Model;

import IEnumMember = powerbi.IEnumMember;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import { LegendData } from "powerbi-visuals-utils-chartutils/lib/legend/legendInterfaces";

export const enum TornadoObjectNames {
    Legend = "legend",
    LegendTitle = "legendTitle",
    Categories = "categories",
    DataPoint = "dataPoint",
    Labels = "labels",
}

class BarsColorGroup extends formattingSettings.Group {
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Fill",
        displayNameKey: "Visual_Fill",
        value: { value: "#000000" }
    });

    name: string = "color";
    displayName: string = "Color";
    displayNameKey: string = "Visual_Color";
    slices = [this.fill];
}

class BarsBorderGroup extends formattingSettings.Group {
    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "" }
    });

    borderWidth = new formattingSettings.Slider({
        name: "borderWidth",
        displayName: "Width (px)",
        displayNameKey: "Visual_BorderWidth",
        value: 0,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    cornerRadius = new formattingSettings.Slider({
        name: "cornerRadius",
        displayName: "Rounded corners (px)",
        displayNameKey: "Visual_CornerRadius",
        value: 4,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 50,
            }
        }
    });

    name: string = "border";
    displayName: string = "Border";
    displayNameKey: string = "Visual_Border";
    slices = [this.borderColor, this.borderWidth, this.cornerRadius];
}

class BarsLayoutGroup extends formattingSettings.Group {
    barSpacing = new formattingSettings.Slider({
        name: "barSpacing",
        displayName: "Space between bars (%)",
        displayNameKey: "Visual_BarSpacing",
        value: 40,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 75,
            }
        }
    });

    name: string = "layout";
    displayName: string = "Layout";
    displayNameKey: string = "Visual_Layout";
    slices = [this.barSpacing];
}

class BarsCardSettings extends CompositeCard {
    color = new BarsColorGroup();
    border = new BarsBorderGroup();
    layout = new BarsLayoutGroup();

    name: string = TornadoObjectNames.DataPoint;
    displayName: string = "Bars";
    displayNameKey: string = "Visual_Bars";
    groups = [this.color, this.border, this.layout];
}

class NegativeBarsColorGroup extends formattingSettings.Group {
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Fill",
        displayNameKey: "Visual_Fill",
        value: { value: "#FFFFFF" }  // white by default
    });

    transparency = new formattingSettings.Slider({
        name: "transparency",
        displayName: "Transparency (%)",
        displayNameKey: "Visual_Transparency",
        value: 0,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 100,
            }
        }
    });

    name: string = "negativeColor";
    displayName: string = "Color";
    displayNameKey: string = "Visual_Color";
    slices = [this.fill, this.transparency];
}

class NegativeBarsBorderGroup extends formattingSettings.Group {
    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "#D64550" }  // red to match negative data labels
    });

    borderWidth = new formattingSettings.Slider({
        name: "borderWidth",
        displayName: "Width (px)",
        displayNameKey: "Visual_BorderWidth",
        value: 1,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    cornerRadius = new formattingSettings.Slider({
        name: "cornerRadius",
        displayName: "Rounded corners (px)",
        displayNameKey: "Visual_CornerRadius",
        value: 4,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 50,
            }
        }
    });

    name: string = "negativeBorder";
    displayName: string = "Border";
    displayNameKey: string = "Visual_Border";
    slices = [this.borderColor, this.borderWidth, this.cornerRadius];
}

class NegativeBarsCardSettings extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true
    });

    topLevelSlice? = this.show;

    color = new NegativeBarsColorGroup();
    border = new NegativeBarsBorderGroup();

    name: string = "negativeBars";
    displayName: string = "Negative bars";
    displayNameKey: string = "Visual_NegativeBars";
    groups = [this.color, this.border];
}

class CenterLineCardSettings extends Card {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true
    });

    topLevelSlice? = this.show;

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "#D3D3D3" }  // light grey
    });

    width = new formattingSettings.Slider({
        name: "width",
        displayName: "Width (px)",
        displayNameKey: "Visual_Width",
        value: 1,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 1,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    name: string = "centerLine";
    displayName: string = "Center line";
    displayNameKey: string = "Visual_CenterLine";
    slices = [this.color, this.width];
}

class PlotAreaCardSettings extends Card {
    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background color",
        displayNameKey: "Visual_BackgroundColor",
        value: { value: null }
    });

    name: string = "plotArea";
    displayName: string = "Plot area background";
    displayNameKey: string = "Visual_PlotArea";
    slices = [this.backgroundColor];
}

class AxisCardSettings extends Card {
    normalize = new formattingSettings.ToggleSwitch({
        name: "normalize",
        displayName: "Normalize to 100%",
        displayNameKey: "Visual_Axis_Normalize",
        value: false
    });

    leftAutoRange = new formattingSettings.ToggleSwitch({
        name: "leftAutoRange",
        displayName: "Left auto range",
        displayNameKey: "Visual_Axis_LeftAutoRange",
        value: true
    });

    leftMin = new formattingSettings.NumUpDown({
        name: "leftMin",
        displayName: "Left minimum",
        displayNameKey: "Visual_Axis_LeftMin",
        value: null
    });

    leftMax = new formattingSettings.NumUpDown({
        name: "leftMax",
        displayName: "Left maximum",
        displayNameKey: "Visual_Axis_LeftMax",
        value: null
    });

    rightAutoRange = new formattingSettings.ToggleSwitch({
        name: "rightAutoRange",
        displayName: "Right auto range",
        displayNameKey: "Visual_Axis_RightAutoRange",
        value: true
    });

    rightMin = new formattingSettings.NumUpDown({
        name: "rightMin",
        displayName: "Right minimum",
        displayNameKey: "Visual_Axis_RightMin",
        value: null
    });

    rightMax = new formattingSettings.NumUpDown({
        name: "rightMax",
        displayName: "Right maximum",
        displayNameKey: "Visual_Axis_RightMax",
        value: null
    });

    name: string = "axis";
    displayName: string = "Axis";
    displayNameKey: string = "Visual_Axis";
    slices = [this.normalize, this.leftAutoRange, this.leftMin, this.leftMax, this.rightAutoRange, this.rightMin, this.rightMax];
}

const labelDisplayFormatOptions: IEnumMember[] = [
    { value: "value", displayName: "Value" },
    { value: "percentage", displayName: "%" },
    { value: "both", displayName: "Value (%)" }
];

const labelPositionOptions: IEnumMember[] = [
    { value: "auto", displayName: "Auto" },
    { value: "insideEnd", displayName: "Inside end" },
    { value: "outsideEnd", displayName: "Outside end" },
    { value: "insideCenter", displayName: "Inside center" },
    { value: "insideBase", displayName: "Inside base" }
];

export class DataLabelSettings extends Card {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true,
    });
    
    topLevelSlice? = this.show;

    displayFormat = new formattingSettings.ItemDropdown({
        name: "displayFormat",
        displayName: "Display format",
        displayNameKey: "Visual_DataLabels_DisplayFormat",
        items: labelDisplayFormatOptions,
        value: labelDisplayFormatOptions[0]
    });

    labelPosition = new formattingSettings.ItemDropdown({
        name: "labelPosition",
        displayName: "Position",
        displayNameKey: "Visual_DataLabels_Position",
        items: labelPositionOptions,
        value: labelPositionOptions[0]
    });

    font: formattingSettings.FontControl = new BaseFontControlSettings(9);

    labelPrecision = new formattingSettings.NumUpDown({
        name: "labelPrecision",
        displayName: "Decimal places",
        displayNameKey: "Visual_DataLabels_DecimalPlaces",
        value: 0,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 25,
            }
        }
    });

    percentagePrecision = new formattingSettings.NumUpDown({
        name: "percentagePrecision",
        displayName: "Percentage decimal places",
        displayNameKey: "Visual_DataLabels_PercentageDecimalPlaces",
        value: 0,
        options: {
            minValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Min,
                value: 0,
            },
            maxValue: {
                type: powerbiVisualsApi.visuals.ValidatorType.Max,
                value: 10,
            }
        }
    });

    labelDisplayUnits = new formattingSettings.AutoDropdown({
        name: "labelDisplayUnits",
        displayName: "Display units",
        displayNameKey: "Visual_DisplayUnits",
        value: 1
    });

    insideFill = new formattingSettings.ColorPicker({
        name: "insideFill",
        displayName: "Inside fill",
        displayNameKey: "Visual_DataLabels_InsideFill",
        value: { value: null }  // null = inherit from theme background
    });

    outsideFill = new formattingSettings.ColorPicker({
        name: "outsideFill",
        displayName: "Outside fill",
        displayNameKey: "Visual_DataLabels_OutsideFill",
        value: { value: "#707070" }
    });

    negativeFill = new formattingSettings.ColorPicker({
        name: "negativeFill",
        displayName: "Negative fill",
        displayNameKey: "Visual_DataLabels_NegativeFill",
        value: { value: "#D64550" }
    });

    name: string = "labels";
    displayName: string = "Data labels";
    displayNameKey: string = "Visual_DataLabels";
    slices = [this.displayFormat, this.labelPosition, this.font, this.labelDisplayUnits, this.labelPrecision, this.percentagePrecision, this.insideFill, this.outsideFill, this.negativeFill];
}

interface IEnumMemberWithDisplayNameKey extends IEnumMember{
    key: string;
}

const positionOptions : IEnumMemberWithDisplayNameKey[] = [
    {value : LegendPosition[LegendPosition.Top], displayName : "Top", key: "Visual_Legend_Position_Top"}, 
    {value : LegendPosition[LegendPosition.Bottom], displayName : "Bottom", key: "Visual_Legend_Position_Bottom"},
    {value : LegendPosition[LegendPosition.Left], displayName : "Left", key: "Visual_Legend_Position_Left"}, 
    {value : LegendPosition[LegendPosition.Right], displayName : "Right", key: "Visual_Legend_Position_Right"}, 
    {value : LegendPosition[LegendPosition.TopCenter], displayName : "Top center", key: "Visual_Legend_Position_Top_Center"}, 
    {value : LegendPosition[LegendPosition.BottomCenter], displayName : "Bottom center", key: "Visual_Legend_Position_Bottom_Center"}, 
    {value : LegendPosition[LegendPosition.LeftCenter], displayName : "Left center", key: "Visual_Legend_Position_Left_Center"}, 
    {value : LegendPosition[LegendPosition.RightCenter], displayName : "Right center", key: "Visual_Legend_Position_Right_Center"}, 
];

class BaseFontCardSettings extends formattingSettings.FontControl {
    private static fontFamilyName: string = "fontFamily";
    private static fontSizeName: string = "fontSize";
    private static boldName: string = "fontBold";
    private static italicName: string = "fontItalic";
    private static underlineName: string = "fontUnderline";
    private static fontName: string = "font";
    public static defaultFontFamily: string = "wf_standard-font, helvetica, arial, sans-serif";
    public static minFontSize: number = 8;
    public static maxFontSize: number = 60;
    constructor(defaultFontSize: number, fontFamily: string = BaseFontCardSettings.defaultFontFamily, settingName: string = ""){
        super(
            new formattingSettings.FontControl({
                name: BaseFontCardSettings.fontName + settingName,
                displayNameKey: "Visual_FontControl",
                fontFamily: new formattingSettings.FontPicker({
                    name: BaseFontCardSettings.fontFamilyName + settingName,
                    value: fontFamily
                }),
                fontSize: new formattingSettings.NumUpDown({
                    name: BaseFontCardSettings.fontSizeName + settingName,
                    displayNameKey: "Visual_FontSize",
                    value: defaultFontSize,
                    options: {
                        minValue: {
                            type: powerbi.visuals.ValidatorType.Min,
                            value: BaseFontCardSettings.minFontSize
                        },
                        maxValue: {
                            type: powerbi.visuals.ValidatorType.Max,
                            value: BaseFontCardSettings.maxFontSize
                        }
                    }
                }),
                bold: new formattingSettings.ToggleSwitch({
                    name: BaseFontCardSettings.boldName + settingName,
                    value: false
                }),
                italic: new formattingSettings.ToggleSwitch({
                    name: BaseFontCardSettings.italicName + settingName,
                    value: false
                }),
                underline: new formattingSettings.ToggleSwitch({
                    name: BaseFontCardSettings.underlineName + settingName,
                    value: false
                })
            })
        );
    }
}

class LegendOptionsGroup extends Card {
    public defaultPosition: IEnumMember = positionOptions[0];

    public position = new formattingSettings.ItemDropdown({
        name: "position",
        displayNameKey: "Visual_Position",
        items: positionOptions,
        value: this.defaultPosition,
    });

    name: string = "legendOptions";
    displayName: string = "Options";
    displayNameKey: string = "Visual_Options";
    slices = [this.position];
}

class LegendTextGroup extends Card {
    public defaultFontSize: number = 12;

    public labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayNameKey: "Visual_LabelColor",
        value: { value: "#616161" },
    });

    public font = new BaseFontCardSettings(this.defaultFontSize, "Segoe UI");

    name: string = "legendText";
    displayName: string = "Text";
    displayNameKey: string = "Visual_Text";
    slices = [this.font, this.labelColor];
}

class LegendTitleGroup extends Card {
    public defaultShowTitle: boolean = false;
    public defaultTitleText: string = "Legend";

    public showTitle = new formattingSettings.ToggleSwitch({
        name: "showTitle",
        displayNameKey: "Visual_ShowTitle",
        value: this.defaultShowTitle,
    });

    topLevelSlice = this.showTitle;

    public titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayNameKey: "Visual_TitleText",
        value: this.defaultTitleText,
        placeholder: "Title text",
    });

    name: string = TornadoObjectNames.LegendTitle;
    displayName: string = "Title";
    displayNameKey: string = "Visual_Title";
    slices = [this.titleText];
}

export class LegendCardSettings extends CompositeCard {
    public defaultShow: boolean = true;

    public name: string = "legend";
    public displayNameKey: string = "Visual_Legend";
    public analyticsPane: boolean = false;

    public show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayNameKey: "Visual_Legend_Show",
        value: this.defaultShow,
    });

    public topLevelSlice: formattingSettings.ToggleSwitch = this.show;

    public options: LegendOptionsGroup = new LegendOptionsGroup();
    public text: LegendTextGroup = new LegendTextGroup();
    public title: LegendTitleGroup = new LegendTitleGroup();

    public groups = [this.options, this.text, this.title];
}

const categoryPositionOptions : IEnumMemberWithDisplayNameKey[] = [
    {value : LegendPosition[LegendPosition.Left], displayName : "Left", key: "Visual_Group_Left"}, 
    {value : LegendPosition[LegendPosition.Right], displayName : "Right", key: "Visual_Group_Right"},
     
];

export class FontDefaultOptions {
    public static DefaultFontSizePt: number = 8;
    public static DefaultFontFamily: string = "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif";
}

export class BaseFontControlSettings extends formattingSettings.FontControl {
    constructor(defaultFontSize: number){
        super(
            new formattingSettings.FontControl({
                name: "font",
                fontFamily: new formattingSettings.FontPicker({
                    name: "fontFamily",
                    value: FontDefaultOptions.DefaultFontFamily
                }),
                fontSize: new formattingSettings.NumUpDown({
                    name: "fontSize",
                    displayName: "Text size",
                    displayNameKey: "Visual_TextSize",
                    value: defaultFontSize,
                    options: {
                        minValue: {
                            type: powerbiVisualsApi.visuals.ValidatorType.Min,
                            value: 8,
                        },
                        maxValue: {
                            type: powerbiVisualsApi.visuals.ValidatorType.Max,
                            value: 60,
                        }
                    }
                }),
                bold: new formattingSettings.ToggleSwitch({
                    name: "fontBold",
                    value: false
                }),
                italic: new formattingSettings.ToggleSwitch({
                    name: "fontItalic",
                    value: false
                }),
                underline: new formattingSettings.ToggleSwitch({
                    name: "fontUnderline",
                    value: false
                })
            })
        );
    }
}

export class CategoryCardSettings extends Card {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show",
        displayNameKey: "Visual_Show",
        value: true,
    });
    
    topLevelSlice? = this.show;

    font: formattingSettings.FontControl = new BaseFontControlSettings(10.5);

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "Color",
        displayNameKey: "Visual_Color",
        value: { value: "#707070" }
    });

    positionDropdown = new formattingSettings.ItemDropdown({
        items: categoryPositionOptions,
        value: categoryPositionOptions[0],
        name: "position",
        displayName: "Position",
        displayNameKey: "Visual_Position"
    });

    name: string = TornadoObjectNames.Categories;
    displayName: string = "Group";
    displayNameKey: string = "Visual_Group";
    slices = [this.positionDropdown, this.font, this.fill];
}


export class TornadoChartSettingsModel extends Model {
    bars = new BarsCardSettings();
    negativeBars = new NegativeBarsCardSettings();
    centerLine = new CenterLineCardSettings();
    plotArea = new PlotAreaCardSettings();
    axis = new AxisCardSettings();
    dataLabels = new DataLabelSettings();
    legend = new LegendCardSettings();
    category = new CategoryCardSettings();

    cards = [
        this.bars,
        this.negativeBars,
        this.centerLine,
        this.axis,
        this.dataLabels,
        this.legend,
        this.category,
        this.plotArea
    ];

    setLocalizedOptions(localizationManager: ILocalizationManager) {
        this.setLocalizedDisplayName(positionOptions, localizationManager);
        this.setLocalizedDisplayName(categoryPositionOptions, localizationManager);
    }   

    public setLocalizedDisplayName(options: IEnumMemberWithDisplayNameKey[], localizationManager: ILocalizationManager) {
        options.forEach(option => {
            option.displayName = localizationManager.getDisplayName(option.key)
        });
    }

    public setVisibilityOfLegendCardSettings(legend: LegendData){
        this.legend.visible = legend.dataPoints.length > 0;
    }
    
    public populateDataColorSlice(dataPoints: TornadoChartSeries[]){
        // Add per-series color pickers to the color group
        const colorSlices: formattingSettings.ColorPicker[] = [];
        for (const dataPoint of dataPoints) {
            // Use series-specific selector for per-series colors
            const seriesSelector = ColorHelper.normalizeSelector(
                dataPoint.selectionId.getSelector(),
                false
            );
            colorSlices.push(
                new formattingSettings.ColorPicker(
                {
                    name: "fill",
                    displayName: dataPoint.name,
                    selector: seriesSelector,
                    value: { value: dataPoint.fill }
                })
            );
        }
        // Update the color group slices with per-series color pickers
        this.bars.color.slices = colorSlices;
    }
}
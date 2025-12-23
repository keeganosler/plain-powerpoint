import * as React from "react";
import { makeStyles, Text, tokens } from "@fluentui/react-components";
import { Checkbox, RadioGroup, Radio } from "@fluentui/react-components";
import { Button } from "@fluentui/react-components";
import "./App.css";
import { ChevronLeft24Regular, ChevronRight24Regular } from "@fluentui/react-icons";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
  },
  footer: {
    marginTop: "auto",
    textAlign: "center",
    backgroundColor: "#f3f3f3",
    paddingTop: `${tokens.spacingVerticalM}`,
  },
  paddedSection: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalXS}`,
  },
  spacedSection: {
    margin: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
  },
});

const App: React.FC<AppProps> = () => {
  const radioOptions = [
    { label: "Apply to entire presentation", value: "all-slides" },
    { label: "Apply to current slide", value: "current-slide" },
  ];
  const [checkboxOptions, setCheckboxOptions] = React.useState([
    { label: "Background images and fills", checked: true },
    { label: "Colors (convert to black & white)", checked: true },
    { label: "Animations and transitions", checked: true },
    { label: "Graphics and decorative images", checked: true },
  ]);  
  const styles = useStyles();
  const [selectedOption, setSelectedOption] = React.useState(radioOptions[0].value);
  const [currentSlide, setCurrentSlide] = React.useState(1);

  const handleRadioChange = (event: React.ChangeEvent<HTMLInputElement>, data: { value: string }) => {
    console.log(event)
    setSelectedOption(data.value);
  }; 

  const handleCheckboxChange = (index: number) => {
    setCheckboxOptions((prevOptions) =>
      prevOptions.map((option, i) =>
        i === index ? { ...option, checked: !option.checked } : option
      )
    );
  };

  const handlePreviousSlide = () => {
    setCurrentSlide((prev) => Math.max(prev - 1, 1));
  };

  const handleNextSlide = () => {
    setCurrentSlide((prev) => Math.min(prev + 1, 12));
  };
  
  const handleSimplifySlides = async () => {
    try {
      await PowerPoint.run(async (context) => {
        const slidesScope =
          selectedOption === "all-slides"
            ? context.presentation.slides
            : context.presentation.getSelectedSlides();

        slidesScope.load("items");
        await context.sync();

        slidesScope.items.forEach((s: any) => {
          s.load("id,index");
          s.shapes.load("items/id,isDecorative,textFrame/textRange/text,background/fill");
        });
        await context.sync();

        if (!slidesScope.items || slidesScope.items.length === 0) {
          console.log("No slides to simplify.");
          return;
        }

        for (const slide of slidesScope.items) {
          const slideBackground = slide.background.fill;
          slideBackground.setSolidFill({color: "#FFFFFF"});

          const texts = (slide.shapes.items || [])
            .filter((sh: any) => !sh.isDecorative)
            .map((sh: any) => (sh.textFrame?.textRange?.text ?? "").trim())
            .filter((t: string) => t.length > 0);

          const aggregated = texts.join("\n\n") || "(no text)";

          const shapeIds = (slide.shapes.items || []).map((sh: any) => sh.id);
          for (const sid of shapeIds) {
            try {
              const s = slide.shapes.getItem(sid);
              s.delete();
            } catch (e) {
              // ignore delete errors for shapes that can't be removed
            }
          }
          await context.sync();

          const textbox = slide.shapes.addTextBox(aggregated, { left: 40, top: 40, width: 860, height: 520 });
          textbox.textFrame.textRange.font.color = "#000000";
          textbox.textFrame.textRange.font.name = "Calibri";
          textbox.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
          await context.sync();
        }
      });
    } catch (err) {
      console.error("Error simplifying slides:", err);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.paddedSection}>
        <div className={`${styles.spacedSection} mb-16`}>
          <Text size={400} weight="semibold">
            Elements to remove:
          </Text>
          <div className="d-flex-col">
            {checkboxOptions.map((option, index) => (
              <Checkbox
                key={index}
                label={option.label}
                checked={option.checked}
                onChange={() => handleCheckboxChange(index)}
              />
            ))}
          </div>          
        </div>
        <div className={`${styles.spacedSection} mb-16`}>
          <Text size={400} weight="semibold">
            Processing Mode:
          </Text>
          <div className="d-flex-col">
            <RadioGroup onChange={handleRadioChange}>
              {radioOptions.map((option, index) => (
                <Radio key={index} label={option.label} value={option.value} />
              ))}
            </RadioGroup>
          </div>          
        </div>
        
        {selectedOption === "current-slide" && (
          <div className="d-flex-row align-items-center justify-content-center mt-16">
            <Button
              icon={<ChevronLeft24Regular />}
              appearance="outline"
              onClick={handlePreviousSlide}
              disabled={currentSlide === 1}
            />
            <Text size={400} weight="semibold" className="m-8">
              {`${currentSlide} of 12 slides`}
            </Text>
            <Button
              icon={<ChevronRight24Regular />}
              appearance="outline"
              onClick={handleNextSlide}
              disabled={currentSlide === 12}
            />
          </div>
        )}
        <Button style={{width: '100%', marginTop: '16px'}} appearance="primary" onClick={handleSimplifySlides}>
          {selectedOption === "all-slides" ? "Convert all slides" : "Convert current slide"}
        </Button>
      </div>
      <div className={styles.footer}>
        This tool preserves all text content while removing visual distractions, making your presentation more accessible for screen readers and viewers with visual processing difficulties.
      </div>        
    </div>
  );
};

export default App;

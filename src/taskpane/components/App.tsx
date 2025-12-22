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
        <Button style={{width: '100%', marginTop: '16px'}} appearance="primary">
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
